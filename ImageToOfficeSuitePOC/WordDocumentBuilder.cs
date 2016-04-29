namespace ImageToOfficeSuitePOC
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Wordprocessing;
    using A = DocumentFormat.OpenXml.Drawing;
    using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

    class WordDocumentBuilder : IDisposable
    {
        const int StandardA4Height = 4760;
        const int StandardPpi = 72;
        const int EmuPerInch = 914400;
        const int EmuRatioForPageSize = EmuPerInch / StandardPpi / 20;

        private readonly WordprocessingDocument document;
        private readonly MainDocumentPart mainPart;

        private readonly PageSize pageSize;
        private readonly PageMargin pageMargin;

        /*
         * 
         * [PageSize: with and height in 20th of a point]
         * 
         * The international default letter size is ISO 216 A4 (210x297mm ~ 8.3×11.7in) and expressed as this:
         * <w:pageSize w:w="11906" w:h="16838"/>
         * https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
         */
        public WordDocumentBuilder(string documentPath)
        {
            document = WordprocessingDocument.Create(documentPath, WordprocessingDocumentType.Document);

            mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var size = new PageSize() { Width = 11906, Height = 16838 };
            var margin = new PageMargin() { Top = 200, Right = 200, Bottom = 200, Left = 200 };

            var sectionProps = new SectionProperties();
            sectionProps.Append(size, margin);

            mainPart.Document.Body.Append(sectionProps);

            pageSize = size;
            pageMargin = margin;
        }

        public void AddImageToDocument(string imagePath)
        {
            using (var imageStream = new FileStream(imagePath, FileMode.Open))
            {
                PrettyPrint($"Processing image: {imagePath}", ConsoleColor.Green);
                var images = GetImagesPerPage(imageStream);

                foreach (var bitmap in images)
                {
                    using (var stream = new MemoryStream())
                    {
                        bitmap.Save(stream, ImageFormat.Png);
                        stream.Position = 0;

                        var documentImagePart = mainPart.AddImagePart(ImagePartType.Png);
                        documentImagePart.FeedData(stream);

                        var extent = GetImageExtentsFor(bitmap);
                        PrettyPrint($"Extents: Cx={extent.Cx,5}, Cy={extent.Cy,5}", ConsoleColor.Cyan);

                        var drawingElement = GetDrawingElement(extent, mainPart.GetIdOfPart(documentImagePart), 0);

                        var p = new Paragraph(new Run(drawingElement));
                        mainPart.Document.Body.Append(p);
                    }
                }
            }
        }

        private IEnumerable<Bitmap> GetImagesPerPage(Stream imageStream)
        {
            using (var sourceBitmap = new Bitmap(imageStream))
            {
                var pxImageWidth = sourceBitmap.Width;
                var pxImageHeight = sourceBitmap.Height;
                var pxPageHeight = StandardA4Height;// - pageMargin.Bottom;

                if (pxImageHeight > pxPageHeight)
                {
                    var numSections = Math.Ceiling(pxImageHeight / (float)pxPageHeight);
                    var pxCropHeight = pxPageHeight;

                    for (var i = 0; i < numSections; i++)
                    {
                        var pxCropStart = pxCropHeight * i;

                        if (pxCropStart + pxCropHeight > pxImageHeight)
                        {
                            pxCropHeight = pxImageHeight - pxCropStart;
                        }

                        var cropRectangle = new Rectangle(0, pxCropStart, pxImageWidth, pxCropHeight);

                        Console.WriteLine($"Rectangle: Y={cropRectangle.Y,5}, Height={cropRectangle.Height,5}");

                        var target = sourceBitmap.Clone(cropRectangle, sourceBitmap.PixelFormat);

                        yield return target;
                    }
                }
                else
                {
                    PrettyPrint($"Rectangle: Y=0, Height={sourceBitmap.Height,5}", ConsoleColor.Red);
                    yield return sourceBitmap;
                }
            }
        }

        private A.Extents GetImageExtentsFor(Bitmap bitmap)
        {
            var xPageMargin = (pageMargin.Left.Value + pageMargin.Right.Value) * EmuRatioForPageSize;
            var yPageMargin = (pageMargin.Top.Value + pageMargin.Bottom.Value) * EmuRatioForPageSize;

            var imageRatio = bitmap.Height / (float)bitmap.Width;

            var emuOriginalWidth = bitmap.Width * (long)(EmuPerInch / bitmap.HorizontalResolution);
            var emuOriginalHeight = bitmap.Height * (long)(EmuPerInch / bitmap.VerticalResolution);

            var emuImageWidth = (long)(pageSize.Width.Value * EmuRatioForPageSize) - xPageMargin;
            var emuImageHeight = (long)(pageSize.Width.Value * imageRatio * EmuRatioForPageSize) - yPageMargin;

            if (emuImageHeight < 0)
            {
                emuImageHeight = emuOriginalHeight;
            }

            // if image is larger than A4 page size, then rescale the image to A4
            // if the image is smaller, use the image's size
            return new A.Extents
            {
                Cx = emuOriginalWidth > emuImageWidth ? emuImageWidth : emuOriginalWidth,
                Cy = emuOriginalHeight > emuImageHeight ? emuImageHeight : emuOriginalHeight
            };
        }

        private static Drawing GetDrawingElement(A.Extents extents, string relationshipId, UInt32Value elementId)
        {
            return new Drawing(new DW.Inline(
             new DW.Extent() { Cx = extents.Cx, Cy = extents.Cy },
             new DW.EffectExtent()
             {
                 LeftEdge = 0L,
                 TopEdge = 0L,
                 RightEdge = 0L,
                 BottomEdge = 0L
             },
             new DW.DocProperties()
             {
                 Id = (UInt32Value)1U,
                 Name = $"Picture {elementId}"
             },
             new DW.NonVisualGraphicFrameDrawingProperties(
                 new A.GraphicFrameLocks() { NoChangeAspect = true }),
             new A.Graphic(
                 new A.GraphicData(
                     new PIC.Picture(
                         new PIC.NonVisualPictureProperties(
                             new PIC.NonVisualDrawingProperties()
                             {
                                 Id = elementId,
                                 Name = $"BitmapImage{elementId}.jpg"
                             },
                             new PIC.NonVisualPictureDrawingProperties()),
                         new PIC.BlipFill(
                             new A.Blip(
                                 new A.BlipExtensionList(
                                     new A.BlipExtension()
                                     {
                                         Uri =
                                           "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                     })
                             )
                             {
                                 Embed = relationshipId,
                                 CompressionState = A.BlipCompressionValues.Print
                             },
                             new A.Stretch(new A.FillRectangle())),
                         new PIC.ShapeProperties(new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, extents),
                             new A.PresetGeometry(new A.AdjustValueList())
                             {
                                 Preset = A.ShapeTypeValues.Rectangle
                             }))
                 )
                 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
         )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });
        }

        public void Dispose()
        {
            this.document.Dispose();
        }

        private static void PrettyPrint(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
