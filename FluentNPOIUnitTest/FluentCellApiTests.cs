using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;

namespace FluentNPOIUnitTest
{
    public class FluentCellApiTests
    {
        [Fact]
        public void SetFormula_ShouldSetFormulaValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act
            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1).SetValue(10)
                .SetCellPosition(ExcelCol.B, 1).SetValue(20)
                .SetCellPosition(ExcelCol.C, 1).SetFormula("A1+B1");

            // Assert
            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(2);
            Assert.NotNull(cell);
            Assert.Equal(CellType.Formula, cell.CellType);
            Assert.Equal("A1+B1", cell.CellFormula);
        }

        [Fact]
        public void SetFormula_WithEqualsPrefix_ShouldRemovePrefix()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetFormula("=SUM(B1:B10)");

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.Equal("SUM(B1:B10)", cell.CellFormula);
        }

        [Fact]
        public void GetPosition_ShouldReturn1BasedRow()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            var cell = fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.B, 5);

            var position = cell.GetPosition();

            Assert.Equal(ExcelCol.B, position.Column);
            Assert.Equal(5, position.Row); // 1-based
        }

        [Fact]
        public void CopyStyleFrom_ShouldCopyStyle()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Set up source cell with style
            fluentWorkbook.SetupCellStyle("SourceStyle", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = IndexedColors.Yellow.Index;
            });

            var sheet = fluentWorkbook.UseSheet("Test");
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("Source")
                .SetCellStyle("SourceStyle");

            // Copy style to another cell
            sheet.SetCellPosition(ExcelCol.B, 1)
                .SetValue("Target")
                .CopyStyleFrom(ExcelCol.A, 1);

            // Assert
            var npSheet = workbook.GetSheet("Test");
            var sourceCell = npSheet.GetRow(0)?.GetCell(0);
            var targetCell = npSheet.GetRow(0)?.GetCell(1);

            Assert.Equal(sourceCell.CellStyle.FillForegroundColor, targetCell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetBackgroundColor_ShouldSetColor()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Colored")
                .SetBackgroundColor(IndexedColors.LightBlue);

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.Equal(IndexedColors.LightBlue.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void SetBorder_ShouldSetAllBorders()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Bordered")
                .SetBorder(BorderStyle.Thin);

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderTop);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderBottom);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderLeft);
            Assert.Equal(BorderStyle.Thin, cell.CellStyle.BorderRight);
        }

        [Fact]
        public void SetAlignment_ShouldSetAlignment()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Centered")
                .SetAlignment(HorizontalAlignment.Center, VerticalAlignment.Top);

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
            Assert.Equal(VerticalAlignment.Top, cell.CellStyle.VerticalAlignment);
        }

        [Fact]
        public void SetNumberFormat_ShouldSetFormat()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue(1234.567)
                .SetNumberFormat("#,##0.00");

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotEqual(0, cell.CellStyle.DataFormat);
        }

        [Fact]
        public void SetWrapText_ShouldEnableWrapText()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Long text that should wrap")
                .SetWrapText(true);

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.True(cell.CellStyle.WrapText);
        }

        [Fact]
        public void SetComment_ShouldAddComment()
        {
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            fluentWorkbook.UseSheet("Test")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("With Comment")
                .SetComment("This is a comment", "Author");

            var sheet = workbook.GetSheet("Test");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell.CellComment);
            Assert.Equal("Author", cell.CellComment.Author);
        }
    }
}
