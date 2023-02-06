using GenericExcelReporter.Core.Contracts;
using GenericExcelReporter.Core.Dtos;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Reflection;

namespace GenericExcelReporter.Core.Services
{
    public class ExcelService : IExcelService
    {
        public async Task<ExcelFileDto> ExportToExcel<T>(List<T> model, string fileName)

        {
            string sWebRootFolder = "Excel";// _hostingEnvironment.WebRootPath;
            Directory.CreateDirectory(sWebRootFolder);
            string sFileName = DateTime.Now.ToString("yy-MM-dd-hh-mm") + $"{fileName}.xlsx";
            var file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));

            var memory = new MemoryStream();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Report");
                worksheet.View.RightToLeft = true;

                //First add the headers
                var titleCells = worksheet.Cells[1, 1];
                int a = 1;
                foreach (var item in typeof(T).GetProperties())
                {
                    var name = item.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;
                    if (name == null)
                    {
                        worksheet.Cells[1, a].Value = item.Name;
                    }
                    else
                    {
                        worksheet.Cells[1, a].Value = name.Name;
                    }

                    a++;
                }
                //worksheet.Cells.DataValidation.AddCustomDataValidation();
                titleCells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                titleCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                if (model.Any())
                {
                    worksheet.Cells["A2"].LoadFromCollection(model);
                    worksheet.Cells["A2"].Style.Numberformat.Format = "@";

                    //worksheet.Cells["A1:A25"].Style.Numberformat.Format = "@";
                }

                ExcelRange range = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                ExcelTable tab = worksheet.Tables.Add(range, "Report");

                tab.TableStyle = TableStyles.Light18;
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                //worksheet.Cells["A1"].LoadFromDataTable(model, true);

                await package.SaveAsync(); //Save the workbook.
            }
            await using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))

            {
                await stream.CopyToAsync(memory);
            }

            memory.Position = 0;
            file.Delete();

            return new ExcelFileDto
            {
                file = memory,

                FileName = sFileName
            };
        }

        public async Task<DataTable> ConvertListToDataTable<T>(List<T> data)
        {
            PropertyDescriptorCollection properties =

            TypeDescriptor.GetProperties(typeof(T));

            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in properties)

                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)

            {
                DataRow row = table.NewRow();

                foreach (PropertyDescriptor prop in properties)

                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;

                table.Rows.Add(row);
            }

            return await Task.FromResult(table);
        }

        public Task<int> GetCountMemberList<T>(List<T> model)
        {
            var countMemberProperties = typeof(T).GetProperties().Length;
            return Task.FromResult(countMemberProperties);
        }
    }
}