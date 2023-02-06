using GenericExcelReporter.Core.Services;
using GenericExcelReporter.Playground.Test.Dtos;

var mockData = new List<Users>()
{
    new Users() { FirstName = "Amirfoad",LastName="Ahmadi",PhoneNumber="09121233434"},
    new Users() { FirstName = "Amirreza",LastName="Ahmadi",PhoneNumber="09121233434"},
};

var excelService = new ExcelService();
var result = excelService.ExportToExcel<Users>(mockData, "Excel").Result;
return System.Text.Encoding.Default.GetBytes(result.file, result.ContentType, result.FileName);