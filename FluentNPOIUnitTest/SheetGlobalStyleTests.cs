using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class SheetGlobalStyleTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; }
        }

        [Fact]
        public void SheetGlobalStyle_ShouldOverrideWorkbookGlobalStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Test1" }
            };

            // Act - Setup workbook-level global style (blue background)
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Blue.Index;
            });

            // Setup Sheet1 with sheet-level global style (green background)
            fluentWorkbook.UseSheet("Sheet1")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Green.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BuildRows();

            // Setup Sheet2 without sheet-level global style (should use workbook global)
            fluentWorkbook.UseSheet("Sheet2")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BuildRows();

            // Assert
            var sheet1 = workbook.GetSheet("Sheet1");
            var sheet2 = workbook.GetSheet("Sheet2");

            var sheet1Cell = sheet1.GetRow(1)?.GetCell(0);
            var sheet2Cell = sheet2.GetRow(1)?.GetCell(0);

            // Sheet1 should have green background (sheet-level global style)
            Assert.NotNull(sheet1Cell?.CellStyle);
            Assert.Equal(FillPattern.SolidForeground, sheet1Cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Green.Index, sheet1Cell.CellStyle.FillForegroundColor);

            // Sheet2 should have blue background (workbook-level global style)
            Assert.NotNull(sheet2Cell?.CellStyle);
            Assert.Equal(FillPattern.SolidForeground, sheet2Cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Blue.Index, sheet2Cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void MultipleSheets_CanHaveDifferentGlobalStyles()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Test1" }
            };

            // Act - Setup different sheet-level global styles for different sheets
            fluentWorkbook.UseSheet("Sheet1")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Red.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BuildRows();

            fluentWorkbook.UseSheet("Sheet2")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Yellow.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BuildRows();

            fluentWorkbook.UseSheet("Sheet3")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Aqua.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BuildRows();

            // Assert
            var sheet1 = workbook.GetSheet("Sheet1");
            var sheet2 = workbook.GetSheet("Sheet2");
            var sheet3 = workbook.GetSheet("Sheet3");

            var sheet1Cell = sheet1.GetRow(1)?.GetCell(0);
            var sheet2Cell = sheet2.GetRow(1)?.GetCell(0);
            var sheet3Cell = sheet3.GetRow(1)?.GetCell(0);

            // Each sheet should have its own global style
            Assert.Equal(IndexedColors.Red.Index, sheet1Cell.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Yellow.Index, sheet2Cell.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Aqua.Index, sheet3Cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SpecificCellStyle_ShouldOverrideSheetGlobalStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Test1" },
                new TestData { ID = 2, Name = "Test2" }
            };

            // Act - Setup sheet-level global style
            fluentWorkbook
                .SetupCellStyle("SpecificStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Orange.Index;
                })
                .UseSheet("Sheet1")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Green.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").SetCellStyle("SpecificStyle").End() // Row 1 with specific style
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End() // Row 2 with sheet global style
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var idCell = sheet.GetRow(1)?.GetCell(0);
            var nameCell = sheet.GetRow(1)?.GetCell(1);

            // ID column should have orange background (specific style overrides global)
            Assert.Equal(IndexedColors.Orange.Index, idCell.CellStyle.FillForegroundColor);

            // Name column should have green background (sheet global style)
            Assert.Equal(IndexedColors.Green.Index, nameCell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SheetGlobalStyle_CanBeUpdated()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Test1" }
            };

            // Act - Setup initial sheet-level global style, then update it
            var sheet = fluentWorkbook.UseSheet("Sheet1")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Red.Index;
                })
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BuildRows();

            // Update the sheet global style
            fluentWorkbook.UseSheet("Sheet1")
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Violet.Index;
                })
                .SetTable(testData, ExcelCol.B, 1)
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BuildRows();

            // Assert
            var sheet1 = workbook.GetSheet("Sheet1");
            var firstCell = sheet1.GetRow(1)?.GetCell(0); // Column A (red style)
            var secondCell = sheet1.GetRow(1)?.GetCell(1); // Column B (violet style)

            // First cell should still have red (created before update)
            Assert.Equal(IndexedColors.Red.Index, firstCell.CellStyle.FillForegroundColor);

            // Second cell should have violet (created after update)
            Assert.Equal(IndexedColors.Violet.Index, secondCell.CellStyle.FillForegroundColor);
        }
    }
}
