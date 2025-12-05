using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class TableTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public bool IsActive { get; set; }
        }

        [Fact]
        public void SetTable_ShouldCreateTableWithHeaders()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", IsActive = true },
                new TestData { ID = 2, Name = "Bob", IsActive = false }
            };

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("編號")
                .BeginBodySet("ID").End()
                .BeginTitleSet("姓名")
                .BeginBodySet("Name").End()
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var titleRow = sheet.GetRow(0);
            Assert.NotNull(titleRow);
            Assert.Equal("編號", titleRow.GetCell(0)?.StringCellValue);
            Assert.Equal("姓名", titleRow.GetCell(1)?.StringCellValue);
        }

        [Fact]
        public void SetTable_WithMultipleRows_ShouldCreateAllRows()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", IsActive = true },
                new TestData { ID = 2, Name = "Bob", IsActive = false },
                new TestData { ID = 3, Name = "Charlie", IsActive = true }
            };

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("姓名")
                .BeginBodySet("Name").End()
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(4, sheet.PhysicalNumberOfRows); // 1 標題 + 3 數據行
        }
    }
}

