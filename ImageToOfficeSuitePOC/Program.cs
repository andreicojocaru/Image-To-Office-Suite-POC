using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImageToOfficeSuitePOC
{
    class Program
    {
        static void Main(string[] args)
        {
            var interopExcel = new InteropExcelGenerator();
            interopExcel.CreateSheet();
        }
    }
}
