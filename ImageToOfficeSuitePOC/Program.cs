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
            WordDocumentBuilder.BuildDocumentWithImage(wordFilePath, imagePath);
            //PdfDocumentBuilder.BuildDocumentWithImage(pdfFilePath, imagePath);
        }
    }
}
