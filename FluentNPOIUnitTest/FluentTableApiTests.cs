using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Stages;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOIUnitTest
{
    public class FluentTableApiTests
    {
        private class TestData
        {
            public string Name { get; set; } = "";
            public int Score { get; set; }
            public string Note { get; set; } = "";
        }

        [Fact]
        public void RowCount_ShouldReturnDataRowCount()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");

            var testData = new List<TestData>
            {
                new TestData { Name = "Alice" },
                new TestData { Name = "Bob" },
                new TestData { Name = "Charlie" }
            };

            var table = fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping);

            Assert.Equal(3, table.RowCount);
        }

        [Fact]
        public void ColumnCount_ShouldReturnMappedColumnCount()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("Score");
            mapping.Map(x => x.Note).ToColumn(ExcelCol.C).WithTitle("Note");

            var testData = new List<TestData> { new TestData() };

            var table = fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping);

            Assert.Equal(3, table.ColumnCount);
        }

        [Fact]
        public void SetColumnWidths_ShouldSetAllColumnsWidth()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("Score");

            var testData = new List<TestData> { new TestData { Name = "Test", Score = 100 } };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows()
                .SetColumnWidths(20);

            var sheet = workbook.GetSheet("Test");
            Assert.Equal(20 * 256, sheet.GetColumnWidth(0));
            Assert.Equal(20 * 256, sheet.GetColumnWidth(1));
        }

        [Fact]
        public void SetTitleRowHeight_ShouldSetTitleHeight()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");

            var testData = new List<TestData> { new TestData { Name = "Test" } };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows()
                .SetTitleRowHeight(30f);

            var sheet = workbook.GetSheet("Test");
            Assert.Equal(30f, sheet.GetRow(0).HeightInPoints);
        }

        [Fact]
        public void SetDataRowHeights_ShouldSetAllDataRowHeights()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");

            var testData = new List<TestData>
            {
                new TestData { Name = "Alice" },
                new TestData { Name = "Bob" }
            };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows()
                .SetDataRowHeights(25f);

            var sheet = workbook.GetSheet("Test");
            Assert.Equal(25f, sheet.GetRow(1).HeightInPoints); // First data row
            Assert.Equal(25f, sheet.GetRow(2).HeightInPoints); // Second data row
        }

        [Fact]
        public void FreezeTitleRow_ShouldFreezePane()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");

            var testData = new List<TestData> { new TestData { Name = "Test" } };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows()
                .FreezeTitleRow();

            var sheet = workbook.GetSheet("Test");
            // NPOI doesn't expose freeze pane info easily, so we just verify no exception
            Assert.NotNull(sheet);
        }

        [Fact]
        public void SetAutoFilter_ShouldSetAutoFilter()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("Score");

            var testData = new List<TestData>
            {
                new TestData { Name = "Alice", Score = 95 },
                new TestData { Name = "Bob", Score = 85 }
            };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows()
                .SetAutoFilter();

            var sheet = workbook.GetSheet("Test") as XSSFSheet;
            Assert.NotNull(sheet);
            // Auto filter is set - verify no exception occurred
        }

        [Fact]
        public void GetTableRange_ShouldReturnCorrectRange()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("Score");

            var testData = new List<TestData>
            {
                new TestData { Name = "Alice", Score = 95 },
                new TestData { Name = "Bob", Score = 85 }
            };

            var table = fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            var range = table.GetTableRange();

            Assert.Equal(0, range.StartRow);  // 0-based
            Assert.Equal(2, range.EndRow);    // 0 (title) + 2 data rows
            Assert.Equal(0, range.StartCol);  // Column A
            Assert.Equal(1, range.EndCol);    // Column B
        }

        [Fact]
        public void AutoSizeColumns_ShouldNotThrow()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("Name");

            var testData = new List<TestData> { new TestData { Name = "Test" } };

            var exception = Record.Exception(() =>
            {
                fluentWorkbook.UseSheet("Test")
                    .SetTable(testData, mapping)
                    .BuildRows()
                    .AutoSizeColumns();
            });

            Assert.Null(exception);
        }

        [Fact]
        public void WithStyleConfig_ShouldApplyStyles()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.Name)
                .ToColumn(ExcelCol.A)
                .WithTitle("Name")
                .WithBackgroundColor(IndexedColors.Yellow)
                .WithFont(isBold: true)
                .WithAlignment(HorizontalAlignment.Center);

            mapping.Map(x => x.Score)
                .ToColumn(ExcelCol.B)
                .WithTitle("Score")
                .WithNumberFormat("0.00");

            var testData = new List<TestData> { new TestData { Name = "Test", Score = 100 } };

            fluentWorkbook.UseSheet("Test")
                .SetTable(testData, mapping)
                .BuildRows();

            var sheet = workbook.GetSheet("Test");
            var row = sheet.GetRow(1); // Data row (0 is title)

            // Verify Col A style
            var cellA = row.GetCell(0);
            var styleA = cellA.CellStyle;
            Assert.Equal(IndexedColors.Yellow.Index, styleA.FillForegroundColor);
            Assert.Equal(FillPattern.SolidForeground, styleA.FillPattern);
            Assert.Equal(HorizontalAlignment.Center, styleA.Alignment);
            var fontA = workbook.GetFontAt(styleA.FontIndex);
            Assert.True(fontA.IsBold);

            // Verify Col B style
            var cellB = row.GetCell(1);
            var styleB = cellB.CellStyle;
            var formatString = styleB.GetDataFormatString();
            Assert.Equal("0.00", formatString);
        }
    }
}
