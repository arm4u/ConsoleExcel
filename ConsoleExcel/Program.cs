using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel; 
using _Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "file.xlsx";
            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(path);
            Worksheet ws = excel.Worksheets[1];

            Console.WriteLine(ws.Cells[1,1].Value2);

            ws.Cells[4, 7] = "Arman!";
            wb.SaveAs("new2.xlsx");
            wb.Close();

            Console.WriteLine("Done!");
            Console.ReadLine();
        }
    }
}
