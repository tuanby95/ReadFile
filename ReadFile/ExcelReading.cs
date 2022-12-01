using IronXL;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;

namespace ReadFile
{
    public class ExcelReading
    {
        //public void ReadExcelasJson()
        //{
        //    var workBook = new WorkBook("ĐƯỜNG CF THỐNG NHẤT(2).xlsx");
        //    var workSheet = workBook.GetWorkSheet("ĐƯỜNG");
        //    var firstCell = workSheet.FirstFilledCell.AddressString;
        //    var lastCell = workSheet.LastFilledCell.AddressString;
        //    string _workSheetRange = $"{firstCell}:{lastCell}";
        //    Console.WriteLine(_workSheetRange);
        //    var range = workSheet.GetRange(_workSheetRange);
        //    foreach(var cell in range )
        //    {
        //        Console.WriteLine(cell);
        //    }
        //    //return workBook.SaveAs("export.json");
        //}
        public void ReadExcel(string strDoc)
        {
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(strDoc);
            Worksheet worksheet = document.Workbook.Worksheets.ByName("ĐƯỜNG");
            var rowCount = worksheet.Rows.LastFormatedRow;
            var columnCount = worksheet.Columns.LastFormatedColumn;
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Console.WriteLine(worksheet.Cell(i, j));
                }
            }
        }
    }
}
