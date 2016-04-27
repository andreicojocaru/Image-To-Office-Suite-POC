namespace ImageToOfficeSuitePOC
{
    using System.Drawing;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

    static class WordDocumentBuilder
    {
        public static void BuildDocumentWithImage(string document, string imagePath)
        {
            using (var wordprocessingDocument =
                WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
            {
                var mainPart = wordprocessingDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var imagePart = mainPart.AddImagePart(ImagePartType.Png);

                using (var stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), imagePath);
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string imagePath)
        {
            //How OOXML units work: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/

            var pageSize = SetPageSize(wordDoc.MainDocumentPart);
            var pageMargin = SetPageMargin(wordDoc.MainDocumentPart);

            const int StandardPpi = 72;
            const int EmuPerInch = 914400;
            const int EmuRatioForPageSize = EmuPerInch / StandardPpi / 20;

            var xPageMargin = (pageMargin.Left.Value + pageMargin.Right.Value) * EmuRatioForPageSize;
            var yPageMargin = (pageMargin.Top.Value + pageMargin.Bottom.Value) * EmuRatioForPageSize;

            var bitmap = new Bitmap(imagePath);
            var imageRatio = bitmap.Height / (float)bitmap.Width;

            var emuBitmapWidth = bitmap.Width * (long)(EmuPerInch / bitmap.HorizontalResolution);
            var emuBitmapHeight = bitmap.Height * (long)(EmuPerInch / bitmap.VerticalResolution);

            var emuImageWidth = (long)(pageSize.Width.Value * EmuRatioForPageSize) - xPageMargin;
            var emuImageHeight = (long)(pageSize.Width.Value * imageRatio * EmuRatioForPageSize) - yPageMargin;

            // if image is larger than A4 page size, then rescale the image to A4
            // if the image is smaller, use the image's size
            var extents = new A.Extents
            {
                Cx = emuBitmapWidth > emuImageWidth ? emuImageWidth : emuBitmapWidth,
                Cy = emuBitmapHeight > emuImageHeight ? emuImageHeight : emuBitmapWidth,
            };

            bitmap.Dispose();

            // Define the reference of the image.
            var element =
     new Drawing(
         new DW.Inline(
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
                 Name = "Picture 1"
             },
             new DW.NonVisualGraphicFrameDrawingProperties(
                 new A.GraphicFrameLocks() { NoChangeAspect = true }),
             new A.Graphic(
                 new A.GraphicData(
                     new PIC.Picture(
                         new PIC.NonVisualPictureProperties(
                             new PIC.NonVisualDrawingProperties()
                             {
                                 Id = (UInt32Value)0U,
                                 Name = "New Bitmap Image.jpg"
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

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(element)));
        }

        private static PageMargin SetPageMargin(MainDocumentPart documentPart)
        {
            var sectionProps = new SectionProperties();
            var pageMargin = new PageMargin() { Top = 100, Right = 100, Bottom = 100, Left = 100 };
            sectionProps.Append(pageMargin);
            documentPart.Document.Body.Append(sectionProps);

            return pageMargin;
        }

        private static PageSize SetPageSize(MainDocumentPart docPart)
        {
            /*
             The international default letter size is ISO 216 A4 (210x297mm ~ 8.3×11.7in) and expressed as this:
               // pageSize: with and height in 20th of a point
                <w:pgSz w:w="11906" w:h="16838"/>
             */
            var sectionProps = new SectionProperties();
            var pageSize = new PageSize() { Width = 11906, Height = 16838 };

            sectionProps.Append(pageSize);
            docPart.Document.Body.Append(sectionProps);

            return pageSize;
        }
    }
}
