using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class BasicTests
    {
        [Fact]
        public void CreateWorkbook_ShouldReturnValidWorkbook()
        {
            // Arrange & Act
            var fluentWorkbook = new FluentWorkbook(new XSSFWorkbook());
            var workbook = fluentWorkbook.GetWorkbook();

            // Assert
            Assert.NotNull(workbook);
            Assert.IsAssignableFrom<IWorkbook>(workbook);
        }

        [Fact]
        public void UseSheet_ShouldCreateNewSheet()
        {
            // Arrange
            var fluentWorkbook = new FluentWorkbook(new XSSFWorkbook());

            // Act
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // Assert
            Assert.NotNull(sheet);
            Assert.NotNull(sheet.GetSheet());
            Assert.Equal("TestSheet", sheet.GetSheet().SheetName);
        }

        [Fact]
        public void SetCellPosition_ShouldSetValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Hello World");

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal("Hello World", cell.StringCellValue);
        }

        [Fact]
        public void SetColumnWidth_ShouldSetWidth()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetColumnWidth(ExcelCol.A, 30);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var width = sheet.GetColumnWidth(0);
            Assert.Equal(30 * 256, width);
        }

        [Fact]
        public void SetCellPosition_WithNumericValue_ShouldSetNumber()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue(123);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(123.0, cell.NumericCellValue);
        }
    }
}

