namespace ImageToOfficeSuitePOC
{
    using System;
    using System.Reflection;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;

    public class InteropExcelGenerator
    {
        public void CreateSheet()
        {
            var xlApp = new Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            ws.Shapes.AddPicture(@"C:\temp\output.png", MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 50, 3367, 3700);

            //// Select the Excel cells, in the range c1 to c7 in the worksheet.
            //Range aRange = ws.get_Range("C1", "C7");

            //if (aRange == null)
            //{
            //    Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            //}

            //// Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            //Object[] args = new Object[1];
            //args[0] = 6;
            //aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

            //// Change the cells in the C1 to C7 range of the worksheet to the number 8.
            //aRange.Value2 = 8;



            wb.SaveAs("MyFile.xlsx", XlFileFormat.xlWorkbookNormal);
        }
    }
}
