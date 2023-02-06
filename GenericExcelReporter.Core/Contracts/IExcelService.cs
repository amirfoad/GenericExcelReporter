using GenericExcelReporter.Core.Dtos;
using System.Data;

namespace GenericExcelReporter.Core.Contracts
{
    public interface IExcelService
    {
        Task<ExcelFileDto> ExportToExcel<T>(List<T> model, string fileName);

        Task<DataTable> ConvertListToDataTable<T>(List<T> model);

        Task<int> GetCountMemberList<T>(List<T> model);
    }
}