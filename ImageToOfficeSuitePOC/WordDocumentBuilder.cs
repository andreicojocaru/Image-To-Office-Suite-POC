namespace ImageToOfficeSuitePOC
{
    using System;
    using System.Drawing;
    using System.IO;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

    class WordDocumentBuilder : IDisposable
    {
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
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using (var stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            using (var stream = imagePart.GetStream(FileMode.Open))
            {
                var imagePartId = mainPart.GetIdOfPart(imagePart);
                var imageParagraph = GetImageParagraph(imagePartId, stream);

                mainPart.Document.Body.Append(imageParagraph);
            }
        }

        private Paragraph GetImageParagraph(string relationshipId, Stream imageStream)
        {
            var extents = GetImageExtentsFor(imageStream);
            var drawingElement = GetDrawingElement(extents, relationshipId, 0);

            return new Paragraph(new Run(drawingElement));
        }

        private A.Extents GetImageExtentsFor(Stream imageStream)
        {
            var xPageMargin = (pageMargin.Left.Value + pageMargin.Right.Value) * EmuRatioForPageSize;
            var yPageMargin = (pageMargin.Top.Value + pageMargin.Bottom.Value) * EmuRatioForPageSize;

            using (var bitmap = new Bitmap(imageStream))
            {
                var imageRatio = bitmap.Height / (float)bitmap.Width;

                var emuOriginalWidth = bitmap.Width * (long)(EmuPerInch / bitmap.HorizontalResolution);
                var emuOriginalHeight = bitmap.Height * (long)(EmuPerInch / bitmap.VerticalResolution);

                var emuImageWidth = (long)(pageSize.Width.Value * EmuRatioForPageSize) - xPageMargin;
                var emuImageHeight = (long)(pageSize.Width.Value * imageRatio * EmuRatioForPageSize) - yPageMargin;

                // if image is larger than A4 page size, then rescale the image to A4
                // if the image is smaller, use the image's size
                return new A.Extents
                {
                    Cx = emuOriginalWidth > emuImageWidth ? emuImageWidth : emuOriginalWidth,
                    Cy = emuOriginalHeight > emuImageHeight ? emuImageHeight : emuOriginalHeight
                };
            }
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
    }
}
