using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class CellStyleRangeTests
    {
        [Fact]
        public void SetCellStyleRange_WithStringKey_ShouldApplyStyleToRange()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 設置一個預定義樣式
            fluentWorkbook.SetupCellStyle("TestRangeStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.LightBlue.Index;
                style.SetBorderAllStyle(BorderStyle.Thin);
            });

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange("TestRangeStyle", ExcelCol.A, ExcelCol.C, 1, 3);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍內的每個單元格都有正確的樣式
            for (int rowIndex = 0; rowIndex <= 2; rowIndex++) // 1-based 轉 0-based: 1-3 -> 0-2
            {
                var row = npoiSheet.GetRow(rowIndex);
                Assert.NotNull(row);

                for (int colIndex = 0; colIndex <= 2; colIndex++) // A-C -> 0-2
                {
                    var cell = row.GetCell(colIndex);
                    Assert.NotNull(cell);
                    Assert.NotNull(cell.CellStyle);
                    Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                    Assert.Equal(IndexedColors.LightBlue.Index, cell.CellStyle.FillForegroundColor);
                    Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);
                }
            }
        }

        [Fact]
        public void SetCellStyleRange_WithCellStyleConfig_ShouldApplyStyleToRange()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var styleConfig = new CellStyleConfig("DynamicRangeStyle", style =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Yellow.Index;
                style.Alignment = HorizontalAlignment.Center;
            });

            // Act
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange(styleConfig, ExcelCol.B, ExcelCol.D, 2, 4);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍內的每個單元格都有正確的樣式
            for (int rowIndex = 1; rowIndex <= 3; rowIndex++) // 2-4 -> 1-3
            {
                var row = npoiSheet.GetRow(rowIndex);
                Assert.NotNull(row);

                for (int colIndex = 1; colIndex <= 3; colIndex++) // B-D -> 1-3
                {
                    var cell = row.GetCell(colIndex);
                    Assert.NotNull(cell);
                    Assert.NotNull(cell.CellStyle);
                    Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                    Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
                    Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
                }
            }
        }

        [Fact]
        public void SetCellStyleRange_SingleCell_ShouldApplyStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.SetupCellStyle("SingleCellStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Green.Index;
            });

            // Act - 只設置單個單元格 (A1)
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange("SingleCellStyle", ExcelCol.A, ExcelCol.A, 1, 1);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);

            Assert.NotNull(cell);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Green.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetCellStyleRange_WithExistingValues_ShouldPreserveValues()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.SetupCellStyle("PreserveStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Orange.Index;
            });

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // 先設置一些值
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Test1");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(true);

            // Act - 對這些已有值的單元格套用樣式
            sheet.SetCellStyleRange("PreserveStyle", ExcelCol.A, ExcelCol.C, 1, 1);

            // Assert - 驗證值被保留且樣式已套用
            var npoiSheet = workbook.GetSheet("Sheet1");
            var row = npoiSheet.GetRow(0);

            var cellA = row.GetCell(0);
            Assert.Equal("Test1", cellA.StringCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellA.CellStyle.FillForegroundColor);

            var cellB = row.GetCell(1);
            Assert.Equal(123.0, cellB.NumericCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellB.CellStyle.FillForegroundColor);

            var cellC = row.GetCell(2);
            Assert.True(cellC.BooleanCellValue);
            Assert.Equal(IndexedColors.Orange.Index, cellC.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetCellStyleRange_LargeRange_ShouldApplyToAllCells()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var styleConfig = new CellStyleConfig("LargeRangeStyle", style =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            });

            // Act - 設置較大範圍 (A1:E10)
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            sheet.SetCellStyleRange(styleConfig, ExcelCol.A, ExcelCol.E, 1, 10);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // 驗證範圍的四個角落和中心點
            var testPoints = new[]
            {
                (row: 0, col: 0), // A1
				(row: 0, col: 4), // E1
				(row: 9, col: 0), // A10
				(row: 9, col: 4), // E10
				(row: 5, col: 2)  // C6 (中心)
			};

            foreach (var (row, col) in testPoints)
            {
                var cell = npoiSheet.GetRow(row)?.GetCell(col);
                Assert.NotNull(cell);
                Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                Assert.Equal(IndexedColors.Grey25Percent.Index, cell.CellStyle.FillForegroundColor);
            }

            // 驗證總共創建了 10 行（1-10）
            Assert.Equal(10, npoiSheet.PhysicalNumberOfRows);
        }
    }
}

