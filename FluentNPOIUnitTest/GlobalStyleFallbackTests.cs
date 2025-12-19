using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;
using System.Collections.Generic;

namespace FluentNPOIUnitTest
{
    public class GlobalStyleFallbackTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public TestData() { }
            public TestData(int id, string name) { ID = id; Name = name; }
        }

        [Fact]
        public void FluentMapping_WithNoStyle_ShouldApplyWorkbookGlobalStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 設定 workbook 全域樣式
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
            });

            // 定義 mapping，不設定任何樣式
            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name");

            fluentWorkbook.UseSheet("TestSheet", true)
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert - 沒指定樣式的 body cells 應該套用 workbook global style
            var sheet = workbook.GetSheet("TestSheet");
            var dataCell = sheet.GetRow(1)?.GetCell(0); // Row 1 is data row (row 0 is title)

            Assert.NotNull(dataCell);
            Assert.Equal(HorizontalAlignment.Center, dataCell.CellStyle.Alignment);
            Assert.Equal(BorderStyle.Thin, dataCell.CellStyle.BorderTop);
            Assert.Equal(BorderStyle.Thin, dataCell.CellStyle.BorderBottom);
        }

        [Fact]
        public void FluentMapping_WithNoStyle_SheetGlobalShouldTakePriorityOverWorkbookGlobal()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 設定 workbook 全域樣式 (藍色背景)
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightBlue.Index;
            });

            // 定義 mapping，不設定任何樣式
            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name");

            // 設定 sheet 全域樣式 (黃色背景) - 應該優先
            fluentWorkbook.UseSheet("TestSheet", true)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Yellow.Index;
                })
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert - 應該套用 sheet global style (黃色)，而非 workbook global (藍色)
            var sheet = workbook.GetSheet("TestSheet");
            var dataCell = sheet.GetRow(1)?.GetCell(0);

            Assert.NotNull(dataCell);
            Assert.Equal(IndexedColors.Yellow.Index, dataCell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void FluentMapping_WithNamedStyle_ShouldTakePriorityOverGlobalStyles()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 設定 workbook 全域樣式 (藍色背景)
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightBlue.Index;
            });

            // 設定具名樣式 (綠色背景)
            fluentWorkbook.SetupCellStyle("GreenStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightGreen.Index;
            });

            // 定義 mapping，ID 欄位使用 GreenStyle
            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID").WithStyle("GreenStyle");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name"); // 無樣式，應該用 global

            fluentWorkbook.UseSheet("TestSheet", true)
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("TestSheet");
            var idCell = sheet.GetRow(1)?.GetCell(0);    // ID 欄位
            var nameCell = sheet.GetRow(1)?.GetCell(1);  // Name 欄位

            Assert.NotNull(idCell);
            Assert.NotNull(nameCell);

            // ID 欄位應該用綠色 (具名樣式優先)
            Assert.Equal(IndexedColors.LightGreen.Index, idCell.CellStyle.FillForegroundColor);

            // Name 欄位應該用藍色 (workbook global fallback)
            Assert.Equal(IndexedColors.LightBlue.Index, nameCell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void FluentMapping_TitleRow_ShouldAlsoFallbackToWorkbookGlobal()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 設定 workbook 全域樣式
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.BorderTop = BorderStyle.Medium;
                style.BorderBottom = BorderStyle.Medium;
            });

            // 定義 mapping，title 沒有設定樣式
            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID"); // 沒有 WithTitleStyle
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name");

            fluentWorkbook.UseSheet("TestSheet", true)
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert - title row 也應該套用 workbook global style
            var sheet = workbook.GetSheet("TestSheet");
            var titleCell = sheet.GetRow(0)?.GetCell(0); // Row 0 is title row

            Assert.NotNull(titleCell);
            Assert.Equal(HorizontalAlignment.Center, titleCell.CellStyle.Alignment);
            Assert.Equal(BorderStyle.Medium, titleCell.CellStyle.BorderTop);
        }

        [Fact]
        public void FluentMapping_MultipleSheets_EachSheetCanHaveOwnGlobalStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 設定 workbook 全域樣式 (白色背景的預設)
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Left;
            });

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");

            // Sheet1 用綠色
            fluentWorkbook.UseSheet("Sheet1", true)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.LightGreen.Index;
                })
                .SetTable(testData, mapping)
                .BuildRows();

            // Sheet2 用黃色
            fluentWorkbook.UseSheet("Sheet2", true)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Yellow.Index;
                })
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert
            var sheet1 = workbook.GetSheet("Sheet1");
            var sheet2 = workbook.GetSheet("Sheet2");

            var cell1 = sheet1.GetRow(1)?.GetCell(0);
            var cell2 = sheet2.GetRow(1)?.GetCell(0);

            Assert.NotNull(cell1);
            Assert.NotNull(cell2);

            Assert.Equal(IndexedColors.LightGreen.Index, cell1.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Yellow.Index, cell2.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void FluentMapping_SheetWithoutSheetGlobal_ShouldUseWorkbookGlobal()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData> { new TestData(1, "Test") };

            // Act - 只設定 workbook 全域樣式
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightBlue.Index;
                style.Alignment = HorizontalAlignment.Center;
            });

            var mapping = new FluentMapping<TestData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");

            // 不設定 sheet 全域樣式
            fluentWorkbook.UseSheet("TestSheet", true)
                .SetTable(testData, mapping)
                .BuildRows();

            // Assert - 應該使用 workbook global style
            var sheet = workbook.GetSheet("TestSheet");
            var dataCell = sheet.GetRow(1)?.GetCell(0);

            Assert.NotNull(dataCell);
            Assert.Equal(IndexedColors.LightBlue.Index, dataCell.CellStyle.FillForegroundColor);
            Assert.Equal(HorizontalAlignment.Center, dataCell.CellStyle.Alignment);
        }
    }
}
