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
                documentBuilder.AddImageToDocument(@"C:\Temp\8k_big_image.png");
                //documentBuilder.AddImageToDocument(@"C:\Temp\output.png");
                //documentBuilder.AddImageToDocument(@"C:\Temp\Bar Chart Page_latest.png");
                //documentBuilder.AddImageToDocument(@"C:\Temp\output.png");
            }
        }
    }
}
