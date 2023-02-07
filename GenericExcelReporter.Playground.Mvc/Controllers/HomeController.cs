using GenericExcelReporter.Core.Contracts;
using GenericExcelReporter.Playground.Mvc.Models;
using Microsoft.AspNetCore.Mvc;

namespace GenericExcelReporter.Playground.Mvc.Controllers
{
    public class HomeController : Controller
    {
        private readonly IExcelService _excelService;

        public HomeController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        private List<UserDto> _mockData
        {
            get
            {
                return new List<UserDto>()

                {
                    new UserDto()
                    { FirstName = "Amirfoad", LastName = "Ahmadi", PhoneNumber = "09121233434",Age=22 },

                    new UserDto()
                    { FirstName = "Amirreza", LastName = "Ahmadi", PhoneNumber = "09121233434",Age=28 },
                };
            }
        }

        public IActionResult Index()
        {
            var result = _excelService.ExportToExcel(_mockData, "Excel").Result;
            return File(result.file, result.ContentType, result.FileName);
        }
    }
}