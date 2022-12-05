// See https://aka.ms/new-console-template for more information
using ReadFile;

string strDoc = @"C:\Users\deheus\My Projects\ReadFile\ReadFile\ĐƯỜNG CF THỐNG NHẤT(2).xlsx";
string filepath = @"C:\Users\deheus\My Projects\ReadFile\ReadFile\75665880_BHOA_GIAO HANG 18.12.2022.docx";
var myExcelReading = new ExcelReading();
var myWordReading = new WordReading();
myWordReading.ReadDocx(filepath);
//myExcelReading.ReadExcel(strDoc);