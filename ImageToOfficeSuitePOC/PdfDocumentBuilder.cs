namespace ImageToOfficeSuitePOC
{
    using PdfSharp.Drawing;
    using PdfSharp.Pdf;

    public static class PdfDocumentBuilder
    {
        public static void BuildDocumentWithImage(string documentPath, string imagePath)
        {
            PdfDocument document = new PdfDocument();

            // Create an empty page or load existing
            PdfPage page = document.AddPage();
            XImage image = XImage.FromFile(imagePath);

            page.Width = image.Size.Width;
            page.Height = image.Size.Height;

            XGraphics gfx = XGraphics.FromPdfPage(page);
            gfx.DrawImage(image, 0, 0, image.Size.Width, image.Size.Height);

            document.Save(documentPath);
        }
    }
}
