using OfficeOpenXml;
using System.Data;

namespace ABS.FileGeneration
{
    public class FileGenerationService
    {
        /// <summary>
        /// Data Table
        /// </summary>
        /// <returns></returns>
        public async Task<DataTable> ExportToExcelAsync()
        {
            DataTable table = new DataTable();
            try
            {
                table.Columns.Add(ColumnClassConstant.Id, typeof(int));
                table.Columns.Add(ColumnClassConstant.Name, typeof(string));
                table.Columns.Add(ColumnClassConstant.Sex, typeof(string));
                table.Columns.Add(ColumnClassConstant.Subject1, typeof(int));
                table.Columns.Add(ColumnClassConstant.Subject2, typeof(int));
                table.Rows.Add(1, "A1", "M", 47, 69);
                table.Rows.Add(2, "A2", "M", 56, 61);
                table.Rows.Add(3, "A3", "F", 69, 63);
                table.Rows.Add(4, "A4", "F", 98, 74);
                table.Rows.Add(5, "A5", "M", 67, 79);
                table.Rows.Add(6, "A6", "M", 52, 89);
            }
            catch (Exception ex) 
            {
                Console.WriteLine($"{ExceptionMessages.ErrorExportToExcel}{ex.Message}");
                throw;
            }
            return await Task.FromResult(table);
        }

        /// <summary>
        /// Generate excel stream
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GenerateFileAsync()
        {
            MemoryStream stream = new MemoryStream();
            List<string> errors = new List<string>();
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(stream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                    var dataTable = await ExportToExcelAsync();
                    workSheet.Cells.LoadFromDataTable(dataTable, true);
                    package.Save();
                }
                stream.Position = 0;

                if(stream.Length == 0)
                {
                    errors.Add(ExceptionMessages.ErrorFileEmpty);
                }
            }
            catch (Exception ex) 
            { 
                errors.Add($"{ExceptionMessages.ErrorGenerateFile}{ex.Message}");
            }
            if (errors.Any())
            {
                throw new AggregateException(ExceptionMessages.ErrorFileGeneration, errors.Select(e => new Exception(e)));
            }
            return stream;
        }
    }
}