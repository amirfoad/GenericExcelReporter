namespace GenericExcelReporter.Core.Dtos
{
    public class ExcelFileDto
    {
        public Stream file { get; set; }
        public string ContentType { get; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public string FileName { get; set; }
    }
}