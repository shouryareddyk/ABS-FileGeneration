using ABS.FileGeneration;
using OfficeOpenXml;
using Xunit;

namespace FileGenerationTest
{
    public class FileGenerationTest
    {
        [Fact]
        public async Task GenerateFile_ReturnsNonEmptyMemoryStream()
        {
            // Arrange
            var fileGenerationService = new FileGenerationService();

            // Act
            var resultStream = await fileGenerationService.GenerateFileAsync();

            // Assert
            Assert.NotNull(resultStream);
            Assert.True(resultStream.Length > 0);
        }

        [Fact]
        public async Task GenerateFile_ProducesValidExcelFile()
        {
            // Arrange
            var fileGenerationService = new FileGenerationService();
            var expectedData = await fileGenerationService.ExportToExcelAsync(); // Await the async method

            // Act
            var resultStream = await fileGenerationService.GenerateFileAsync();

            // Assert
            resultStream.Position = 0; // Reset stream position
            using (var package = new ExcelPackage(resultStream))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                Assert.NotNull(worksheet);

                // Validate data
                // Assuming headers are present, start validating from the second row (Excel rows are 1-based)
                for (int row = 2; row <= expectedData.Rows.Count + 1; row++)
                {
                    for (int col = 1; col <= expectedData.Columns.Count; col++)
                    {
                        var expectedValue = expectedData.Rows[row - 2][col - 1]?.ToString() ?? string.Empty; // Adjust row index for DataTable (0-based)
                        var actualValue = worksheet.Cells[row, col].Text; // Adjust row index to skip header in Excel (1-based)
                        Assert.Equal(expectedValue, actualValue);
                    }
                }
            }
        }


        [Fact]
        public async Task GenerateFile_WorksheetIsNotNull()
        {
            // Arrange
            var fileGenerationService = new FileGenerationService();

            // Act
            var resultStream = await fileGenerationService.GenerateFileAsync();

            // Assert
            resultStream.Position = 0; // Reset stream position
            using (var package = new ExcelPackage(resultStream))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"]; // Retrieve the worksheet from the generated Excel file
                Assert.NotNull(worksheet); // Verify that the retrieved worksheet is not null
            }
        }

        [Fact]
        public async Task GenerateFile_GetResult()
        {
            // Arrange
            var fileGenerationService = new FileGenerationService();

            // Act
            var actualDataTable = await fileGenerationService.ExportToExcelAsync(); // Use the async method correctly

            // Assert
            Assert.Equal(6, actualDataTable.Rows.Count); // Directly checking row count
            Assert.Equal(5, actualDataTable.Columns.Count); // Directly checking Column count
        }
    }
}
