namespace ImageToOfficeSuitePOC
{
    using System;

    using PdfSharp;
    using PdfSharp.Drawing;
    using PdfSharp.Pdf;

    public static class PdfDocumentBuilder
    {
        public static void BuildDocumentWithImage(string documentPath, string imagePath)
        {
            using (var document = new PdfDocument())
            {
                using (var image = XImage.FromFile(imagePath))
                {
                    var margin = 30;
                    var totalHeight = image.PixelHeight / 7.2;

                    var numPages =
                        Math.Ceiling(
                            totalHeight
                            / new PdfPage { Size = PageSize.A4, Orientation = PageOrientation.Portrait }.Height);

                    totalHeight = totalHeight - numPages * margin * 2;

                    var pageIndex = 0;
                    double saveHeight = 0;
                    while (saveHeight < totalHeight)
                    {
                        var page = new PdfPage { Size = PageSize.A4, Orientation = PageOrientation.Portrait };

                        document.Pages.Add(page);

                        var pageWidth = page.Width - margin * 2;
                        var pageHeight = page.Height - margin * 2;

                        XGraphics xgr = XGraphics.FromPdfPage(document.Pages[pageIndex]);
                        xgr.DrawImage(image, margin, margin + (-pageIndex * pageHeight), pageWidth, totalHeight);
                        saveHeight += pageHeight;

                        pageIndex++;
                    }
                }

                document.Save(documentPath);
            }
        }
    }
}


//namespace BankBI.Web.Services.DocumentGenerators.Builders
//{
//    using System;

//    using PdfSharp;
//    using PdfSharp.Drawing;
//    using PdfSharp.Pdf;

//    public static class PdfDocumentBuilder
//    {
//        public static void BuildDocumentWithImage(string documentPath, string imagePath)
//        {
//            using (var document = new PdfDocument())
//            {
//                using (var image = XImage.FromFile(imagePath))
//                {
//                    var margin = 50;

//                    var page = new PdfPage { Size = PageSize.A4, Orientation = PageOrientation.Portrait };
//                    document.Pages.Add(page);

//                    var pageWidth = page.Width - margin * 2;
//                    var pageHeight = image.PointHeight > page.Height ? page.Height - margin : image.PointHeight;

//                    XGraphics xgr = XGraphics.FromPdfPage(document.Pages[0]);
//                    xgr.DrawImage(image, margin, margin, pageWidth, pageHeight);

//                    //var margin = 50;
//                    //var totalHeight = image.PixelHeight / 5.8;

//                    //var numPages = Math.Ceiling(totalHeight / new PdfPage { Size = PageSize.A4, Orientation = PageOrientation.Portrait }.Height);

//                    //totalHeight = totalHeight - numPages * margin * 2;

//                    //var pageIndex = 0;
//                    //double saveHeight = 0;
//                    //while (saveHeight < totalHeight)
//                    //{
//                    //    var page = new PdfPage { Size = PageSize.A4, Orientation = PageOrientation.Portrait };

//                    //    document.Pages.Add(page);

//                    //    var pageWidth = page.Width - margin * 2;
//                    //    var pageHeight = page.Height - margin * 2;

//                    //    XGraphics xgr = XGraphics.FromPdfPage(document.Pages[pageIndex]);
//                    //    xgr.DrawImage(image, margin, margin + (-pageIndex * pageHeight), pageWidth, totalHeight);
//                    //    saveHeight += pageHeight;

//                    //    pageIndex++;
//                    //}
//                }

//                document.Save(documentPath);
//            }
//        }
//    }
//}
