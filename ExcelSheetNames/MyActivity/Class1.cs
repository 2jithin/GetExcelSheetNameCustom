using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace MyActivity
{
    public class ExcelSheetNames : CodeActivity
    {
        [Category("Input")]
        [DisplayName("FilePath")]
        [Description("Enter Full file path")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        [DisplayName("Sheet Names")]
        [Description("The sheet names in a list of string")]
        public OutArgument<List<string>> SheetNames { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            List<string> sheets = new List<string>();
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
                    //Console.WriteLine(worksheet.Name);
                    sheets.Add(worksheet.Name);                    
                }
            }
            var result = sheets; // do your stuff
            SheetNames.Set(context, result);

        }
    }
}
