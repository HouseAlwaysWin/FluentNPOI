using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FluentNPOI.Stages;

namespace FluentNPOIUnitTest
{
    public class StyleInheritanceTests
    {
        [Fact]
        public void SetupCellStyle_WithInheritFrom_ShouldInheritParentProperties()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 設定 global 樣式
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
            });

            // 設定繼承 global 的樣式，只覆寫背景色
            fluentWorkbook.SetupCellStyle("HeaderBlue", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightCornflowerBlue.Index;
            }, inheritFrom: "global");

            // Assert - HeaderBlue 應該繼承 global 的對齊和邊框
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Test")
                .SetCellStyle("HeaderBlue");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment); // 繼承自 global
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);           // 繼承自 global
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderBottom);        // 繼承自 global
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern); // 自己設定
            Assert.Equal(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor); // 自己設定
        }

        [Fact]
        public void SetupCellStyle_WithMultiLevelInheritance_ShouldInheritAllLevels()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 設定 global
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
            });

            // HeaderBase 繼承 global
            fluentWorkbook.SetupCellStyle("HeaderBase", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            }, inheritFrom: "global");

            // HeaderBlue 繼承 HeaderBase，覆寫顏色
            fluentWorkbook.SetupCellStyle("HeaderBlue", (wb, style) =>
            {
                style.FillForegroundColor = IndexedColors.LightCornflowerBlue.Index;
            }, inheritFrom: "HeaderBase");

            // Assert
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Test")
                .SetCellStyle("HeaderBlue");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);        // 來自 global
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);                  // 來自 global
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);     // 來自 HeaderBase
            Assert.Equal(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor); // 自己覆寫
        }

        [Fact]
        public void SetupCellStyle_WithPropertyOverride_ShouldOverrideParentProperty()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 設定 global 為置中對齊
            fluentWorkbook.SetupGlobalCachedCellStyles((wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
            });

            // AmountStyle 繼承 global，但覆寫為靠右對齊
            fluentWorkbook.SetupCellStyle("AmountStyle", (wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Right;
            }, inheritFrom: "global");

            // Assert
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue(123.45)
                .SetCellStyle("AmountStyle");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Right, cell.CellStyle.Alignment); // 應該是覆寫後的值
        }

        [Fact]
        public void SetupCellStyle_WithNonExistentParent_ShouldStillWork()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 嘗試繼承不存在的樣式
            fluentWorkbook.SetupCellStyle("TestStyle", (wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Yellow.Index;
            }, inheritFrom: "NonExistentStyle"); // 不存在的父樣式

            // Assert - 應該正常運作，只是沒有繼承
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Test")
                .SetCellStyle("TestStyle");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
            Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetupCellStyle_WithoutInheritFrom_ShouldWorkAsUsual()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 不使用 inheritFrom
            fluentWorkbook.SetupCellStyle("TestStyle", (wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Left;
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Red.Index;
            });

            // Assert
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Test")
                .SetCellStyle("TestStyle");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Left, cell.CellStyle.Alignment);
            Assert.Equal(IndexedColors.Red.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetupCellStyle_InheritFromOtherNamedStyle_ShouldWork()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 設定一個具名樣式 (不是 global)
            fluentWorkbook.SetupCellStyle("BaseStyle", (wb, style) =>
            {
                style.Alignment = HorizontalAlignment.Center;
                style.BorderTop = BorderStyle.Medium;
                style.BorderBottom = BorderStyle.Medium;
            });

            // 從具名樣式繼承
            fluentWorkbook.SetupCellStyle("DerivedStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Green.Index;
            }, inheritFrom: "BaseStyle");

            // Assert
            fluentWorkbook.UseSheet("TestSheet")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Test")
                .SetCellStyle("DerivedStyle");

            var sheet = workbook.GetSheet("TestSheet");
            var cell = sheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);      // 來自 BaseStyle
            Assert.Equal(BorderStyle.Medium, cell.CellStyle.BorderTop);               // 來自 BaseStyle
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);   // 自己設定
            Assert.Equal(IndexedColors.Green.Index, cell.CellStyle.FillForegroundColor); // 自己設定
        }
    }
}
