using Xunit;
using FluentNPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Regression tests for the cached FluentCell style setters: applying the same
    /// modification to many cells must reuse a single ICellStyle (NPOI caps cell styles
    /// at ~64k per workbook) while still producing the correct visual result.
    /// </summary>
    public class StyleCacheTests
    {
        [Fact]
        public void SetBackgroundColor_AppliedToManyCells_ReusesSingleStyle()
        {
            var workbook = new XSSFWorkbook();
            var fluent = new FluentWorkbook(workbook);
            var sheet = fluent.UseSheet("Test");

            int before = workbook.NumCellStyles;
            for (int r = 1; r <= 500; r++)
            {
                sheet.SetCellPosition(ExcelCol.A, r)
                     .SetValue(r)
                     .SetBackgroundColor(IndexedColors.Yellow);
            }
            int added = workbook.NumCellStyles - before;

            // 500 cells, identical modification from the same base style => one new style.
            Assert.True(added <= 2, $"Expected at most 2 new cell styles, got {added}");
        }

        [Fact]
        public void SetBackgroundColor_AppliesCorrectColorAndPattern()
        {
            var workbook = new XSSFWorkbook();
            var fluent = new FluentWorkbook(workbook);

            fluent.UseSheet("Test")
                  .SetCellPosition(ExcelCol.A, 1)
                  .SetBackgroundColor(IndexedColors.Yellow);

            var cell = workbook.GetSheet("Test").GetRow(0).GetCell(0);
            Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
            Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
        }

        [Fact]
        public void ChainedSetters_ComposeAndStayBounded()
        {
            var workbook = new XSSFWorkbook();
            var fluent = new FluentWorkbook(workbook);
            var sheet = fluent.UseSheet("Test");

            int before = workbook.NumCellStyles;
            for (int r = 1; r <= 200; r++)
            {
                sheet.SetCellPosition(ExcelCol.A, r)
                     .SetValue(r)
                     .SetBackgroundColor(IndexedColors.Red)
                     .SetBorder(BorderStyle.Thin)
                     .SetAlignment(HorizontalAlignment.Center, VerticalAlignment.Center);
            }
            int added = workbook.NumCellStyles - before;

            // A fixed 3-step chain from the same base => a small constant number of styles,
            // independent of cell count.
            Assert.True(added <= 4, $"Expected at most 4 new cell styles, got {added}");

            var cell = workbook.GetSheet("Test").GetRow(0).GetCell(0);
            Assert.Equal(IndexedColors.Red.Index, cell.CellStyle.FillForegroundColor);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderRight);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
        }

        [Fact]
        public void DifferentModifications_ProduceIndependentStyles()
        {
            var workbook = new XSSFWorkbook();
            var fluent = new FluentWorkbook(workbook);
            var sheet = fluent.UseSheet("Test");

            sheet.SetCellPosition(ExcelCol.A, 1).SetBackgroundColor(IndexedColors.Red);
            sheet.SetCellPosition(ExcelCol.B, 1).SetBackgroundColor(IndexedColors.Blue);

            var row = workbook.GetSheet("Test").GetRow(0);
            Assert.Equal(IndexedColors.Red.Index, row.GetCell(0).CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Blue.Index, row.GetCell(1).CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetNumberFormat_AppliedToManyCells_ReusesSingleStyle()
        {
            var workbook = new XSSFWorkbook();
            var fluent = new FluentWorkbook(workbook);
            var sheet = fluent.UseSheet("Test");

            int before = workbook.NumCellStyles;
            for (int r = 1; r <= 300; r++)
            {
                sheet.SetCellPosition(ExcelCol.A, r)
                     .SetValue(r)
                     .SetNumberFormat("#,##0.00");
            }
            int added = workbook.NumCellStyles - before;

            Assert.True(added <= 2, $"Expected at most 2 new cell styles, got {added}");

            var cell = workbook.GetSheet("Test").GetRow(0).GetCell(0);
            Assert.Equal("#,##0.00", cell.CellStyle.GetDataFormatString());
        }
    }
}
