using IronXL;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFile
{
    public class ExcelReading
    {
        public void ReadExcelasJson()
        {
            var workBook = new WorkBook("ĐƯỜNG CF THỐNG NHẤT(2).xlsx");
            var workSheet = workBook.GetWorkSheet("ĐƯỜNG");
            var firstCell = workSheet.FirstFilledCell.AddressString;
            var lastCell = workSheet.LastFilledCell.AddressString;
            string _workSheetRange = $"{firstCell}:{lastCell}";
            Console.WriteLine(_workSheetRange);
            var range = workSheet.GetRange(_workSheetRange);
            foreach(var cell in range )
            {
                Console.WriteLine(cell);
            }
            //return workBook.SaveAs("export.json");
        }
    }
}
