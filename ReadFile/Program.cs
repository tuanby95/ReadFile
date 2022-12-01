// See https://aka.ms/new-console-template for more information
using ReadFile;

string strDoc = @"D:\C#\Project\ReadFile\ĐƯỜNG CF THỐNG NHẤT(2).xlsx";
var myExcelReading = new ExcelReading();
myExcelReading.ReadExcel(strDoc);