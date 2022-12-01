// See https://aka.ms/new-console-template for more information
using ReadFile;

string strDoc = @"D:\C#\Project\ReadFile\testdata.xlsx";
var myExcelReading = new ExcelReading();
myExcelReading.ReadExcelasJson();