using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FluentNPOI.Stages;

namespace FluentNPOIUnitTest
{
    public class CopyStyleFromSheetCellTests
    {
        [Fact]
        public void CopyStyleFromSheetCell_ShouldCacheStyleAtWorkbookLevel()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 創建模板工作表並設置樣式
            var templateSheet = fluentWorkbook
                .SetupCellStyle("templateStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Aqua.Index;
                    style.Alignment = HorizontalAlignment.Right;
                })
                .UseSheet("Template");

            templateSheet.SetCellPosition(ExcelCol.B, 3)
                .SetCellStyle("templateStyle")
                .SetValue("Template Cell");

            // Act - 從模板工作表複製樣式到工作簿層級
            var templateSheetRef = templateSheet.GetSheet();
            fluentWorkbook.CopyStyleFromSheetCell("copiedFromTemplate", templateSheetRef, ExcelCol.B, 3);

            // 在另一個工作表中使用該樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelCol.A, 1)
                .SetCellStyle("copiedFromTemplate")
                .SetValue("Using Copied Style");

            // Assert
            var npoiDataSheet = workbook.GetSheet("Data");
            var cell = npoiDataSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal("Using Copied Style", cell.StringCellValue);
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Aqua.Index, cell.CellStyle.FillForegroundColor);
            Assert.Equal(HorizontalAlignment.Right, cell.CellStyle.Alignment);
        }

        [Fact]
        public void CopyStyleFromSheetCell_ShouldWorkAcrossMultipleSheets()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 在第一個工作表創建樣式
            var sheet1 = fluentWorkbook
                .SetupCellStyle("originalStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.LightGreen.Index;
                    var font = wb.CreateFont();
                    font.IsBold = true;
                    style.SetFont(font);
                })
                .UseSheet("Sheet1");

            sheet1.SetCellPosition(ExcelCol.A, 1)
                .SetCellStyle("originalStyle")
                .SetValue("Source");

            var sheet1Ref = sheet1.GetSheet();

            // Act - 複製樣式並在多個工作表中使用
            fluentWorkbook.CopyStyleFromSheetCell("sharedStyle", sheet1Ref, ExcelCol.A, 1);

            var sheet2 = fluentWorkbook.UseSheet("Sheet2");
            sheet2.SetCellPosition(ExcelCol.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet2 Data");

            var sheet3 = fluentWorkbook.UseSheet("Sheet3");
            sheet3.SetCellPosition(ExcelCol.A, 1).SetCellStyle("sharedStyle").SetValue("Sheet3 Data");

            // Assert - 驗證所有工作表都使用了相同的樣式
            var npoiSheet2 = workbook.GetSheet("Sheet2");
            var npoiSheet3 = workbook.GetSheet("Sheet3");

            var cell2 = npoiSheet2.GetRow(0)?.GetCell(0);
            var cell3 = npoiSheet3.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.Equal(IndexedColors.LightGreen.Index, cell2.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.LightGreen.Index, cell3.CellStyle.FillForegroundColor);
            Assert.True(cell2.CellStyle.GetFont(workbook).IsBold);
            Assert.True(cell3.CellStyle.GetFont(workbook).IsBold);
        }

        [Fact]
        public void CopyStyleFromSheetCell_WithNonExistentCell_ShouldNotCache()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var emptySheet = fluentWorkbook.UseSheet("Empty").GetSheet();

            // Act - 嘗試從不存在的單元格複製樣式
            fluentWorkbook.CopyStyleFromSheetCell("nonExistent", emptySheet, ExcelCol.Z, 999);

            // 嘗試使用該樣式 - 應該不會有特殊樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelCol.A, 1)
                .SetCellStyle("nonExistent")
                .SetValue("Test");

            // Assert
            var npoiSheet = workbook.GetSheet("Test");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            // 因為樣式未被緩存，單元格應該使用默認樣式
            Assert.Equal(FillPattern.NoFill, cell.CellStyle.FillPattern);
        }

        [Fact]
        public void CopyStyleFromSheetCell_SameCacheKeyTwice_ShouldNotOverwrite()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 創建兩個不同樣式的工作表
            var blueSheet = fluentWorkbook
                .SetupCellStyle("blue", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Blue.Index;
                })
                .UseSheet("BluePalette");
            blueSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("blue").SetValue("Blue");

            var redSheet = fluentWorkbook
                .SetupCellStyle("red", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Red.Index;
                })
                .UseSheet("RedPalette");
            redSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("red").SetValue("Red");

            var blueSheetRef = blueSheet.GetSheet();
            var redSheetRef = redSheet.GetSheet();

            // Act - 第一次複製藍色樣式
            fluentWorkbook.CopyStyleFromSheetCell("myColor", blueSheetRef, ExcelCol.A, 1);

            // 第二次嘗試用相同的鍵複製紅色樣式 - 應該被忽略
            fluentWorkbook.CopyStyleFromSheetCell("myColor", redSheetRef, ExcelCol.A, 1);

            // 使用該樣式
            var testSheet = fluentWorkbook.UseSheet("Test");
            testSheet.SetCellPosition(ExcelCol.A, 1)
                .SetCellStyle("myColor")
                .SetValue("Test");

            // Assert - 應該仍然是第一次複製的藍色
            var npoiSheet = workbook.GetSheet("Test");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.Blue.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void CopyStyleFromSheetCell_ChainedCalls_ShouldWork()
        {
            // Arrange \u0026 Act
            var workbook = new XSSFWorkbook();

            // 使用鏈式調用複製多個樣式
            var fluentWorkbook = new FluentWorkbook(workbook);

            var paletteSheet = fluentWorkbook.UseSheet("Palette");

            // 設置多個樣式單元格
            fluentWorkbook
                .SetupCellStyle("style1", (wb, s) => s.FillForegroundColor = IndexedColors.Yellow.Index)
                .SetupCellStyle("style2", (wb, s) => s.FillForegroundColor = IndexedColors.Green.Index)
                .SetupCellStyle("style3", (wb, s) => s.FillForegroundColor = IndexedColors.Orange.Index);

            paletteSheet.SetCellPosition(ExcelCol.A, 1).SetCellStyle("style1").SetValue("Style1");
            paletteSheet.SetCellPosition(ExcelCol.A, 2).SetCellStyle("style2").SetValue("Style2");
            paletteSheet.SetCellPosition(ExcelCol.A, 3).SetCellStyle("style3").SetValue("Style3");

            var paletteSheetRef = paletteSheet.GetSheet();

            // 鏈式調用複製所有樣式
            fluentWorkbook
                .CopyStyleFromSheetCell("copied1", paletteSheetRef, ExcelCol.A, 1)
                .CopyStyleFromSheetCell("copied2", paletteSheetRef, ExcelCol.A, 2)
                .CopyStyleFromSheetCell("copied3", paletteSheetRef, ExcelCol.A, 3);

            // 使用複製的樣式
            var dataSheet = fluentWorkbook.UseSheet("Data");
            dataSheet.SetCellPosition(ExcelCol.B, 1).SetCellStyle("copied1").SetValue("Copy1");
            dataSheet.SetCellPosition(ExcelCol.B, 2).SetCellStyle("copied2").SetValue("Copy2");
            dataSheet.SetCellPosition(ExcelCol.B, 3).SetCellStyle("copied3").SetValue("Copy3");

            // Assert
            var npoiDataSheet = workbook.GetSheet("Data");

            var cell1 = npoiDataSheet.GetRow(0)?.GetCell(1);
            var cell2 = npoiDataSheet.GetRow(1)?.GetCell(1);
            var cell3 = npoiDataSheet.GetRow(2)?.GetCell(1);

            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);

            Assert.Equal(IndexedColors.Yellow.Index, cell1.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Green.Index, cell2.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Orange.Index, cell3.CellStyle.FillForegroundColor);
        }
    }
}

