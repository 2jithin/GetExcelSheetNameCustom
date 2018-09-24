using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace GetSheetNames
{
    class Program
    {
        static void Main(string[] args)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();  // Creates a new Excel Application
            excelApp.Visible = true;  // Makes Excel visible to the user.           
                                      // The following code opens an existing workbook
            string workbookPath = @"C:\Users\akalbhor\Desktop\Salary Slips\ParcelMonkey\Parcel Monkey.xlsx";
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                false, 0, true, false, false);
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                excelWorkbook = excelApp.Workbooks.Add();
            }
            // The following gets the Worksheets collection
            Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            foreach (Worksheet worksheet in excelWorkbook.Worksheets)
            {
                Console.WriteLine(worksheet.Name);
            }
        }
    }
}
