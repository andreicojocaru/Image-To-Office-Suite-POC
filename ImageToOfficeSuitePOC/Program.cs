namespace ImageToOfficeSuitePOC
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            var excelFilePath = @"C:\Temp\demo.xlsx";
            var wordFilePath = @"C:\Temp\demo.docx";
            var pdfFilePath = @"C:\Temp\demo.pdf";

            var imagePath = @"C:\Temp\output.png";

            //ExcelDocumentBuilder.BuildDocumentWithImage(excelFilePath, imagePath);
            //PdfDocumentBuilder.BuildDocumentWithImage(pdfFilePath, imagePath);

            using (var documentBuilder = new WordDocumentBuilder(wordFilePath))
            {
                documentBuilder.AddImageToDocument(@"C:\Temp\3k.png");
                documentBuilder.AddImageToDocument(@"C:\Temp\6k.png");
                documentBuilder.AddImageToDocument(@"C:\Temp\9k.png");
                documentBuilder.AddImageToDocument(@"C:\Temp\14k.png");
            }
        }
    }
}
