using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload.Widgets;
using FluentNPOI.HotReload.Context;
using FluentNPOI.HotReload.Diagnostics;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Widget Engine Core (Phase 1)
    /// </summary>
    public class WidgetRenderingTests
    {
        #region DebugLocation Tests

        [Fact]
        public void DebugLocation_ToComment_ShouldReturnFormattedString()
        {
            // Arrange
            var location = new DebugLocation(@"C:\Projects\Test\MyWidget.cs", 42);

            // Act
            var comment = location.ToComment();

            // Assert
            Assert.Equal("Source: MyWidget.cs:42", comment);
        }

        [Fact]
        public void DebugLocation_ToString_ShouldReturnFullPath()
        {
            // Arrange
            var location = new DebugLocation(@"C:\Projects\Test\MyWidget.cs", 42);

            // Act
            var result = location.ToString();

            // Assert
            Assert.Equal(@"C:\Projects\Test\MyWidget.cs:42", result);
        }

        #endregion

        #region ExcelContext Tests

        [Fact]
        public void ExcelContext_SetValue_ShouldWriteToCurrentPosition()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Act
            ctx.SetValue("Test Value");

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal("Test Value", cell.StringCellValue);
        }

        [Fact]
        public void ExcelContext_MoveToNextRow_ShouldIncrementRowPosition()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Act
            ctx.SetValue("Row 1");
            ctx.MoveToNextRow();
            ctx.SetValue("Row 2");

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Row 1", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Row 2", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
        }

        [Fact]
        public void ExcelContext_MoveToNextColumn_ShouldIncrementColumnPosition()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Act
            ctx.SetValue("Col A");
            ctx.MoveToNextColumn();
            ctx.SetValue("Col B");

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Col A", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Col B", npoiSheet.GetRow(0)?.GetCell(1)?.StringCellValue);
        }

        [Fact]
        public void ExcelContext_MoveTo_ShouldSetExactPosition()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Act
            ctx.MoveTo(ExcelCol.C, 5);
            ctx.SetValue("C5 Value");

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(4)?.GetCell(2); // Row 5 = index 4, Col C = index 2
            Assert.NotNull(cell);
            Assert.Equal("C5 Value", cell.StringCellValue);
        }

        [Fact]
        public void ExcelContext_SetBackgroundColor_ShouldApplyColor()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Act
            ctx.SetValue("Colored");
            ctx.SetBackgroundColor(IndexedColors.Yellow);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
        }

        #endregion

        #region Label Widget Tests

        [Fact]
        public void Label_Build_ShouldSetText()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var label = new Label("Hello Label");

            // Act
            label.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal("Hello Label", cell.StringCellValue);
        }

        [Fact]
        public void Label_SetBold_ShouldApplyBoldStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var label = new Label("Bold Label").SetBold();

            // Act
            label.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            var font = workbook.GetFontAt(cell.CellStyle.FontIndex);
            Assert.True(font.IsBold);
        }

        #endregion

        #region Header Widget Tests

        [Fact]
        public void Header_Build_ShouldSetTextAndBold()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var header = new Header("Title");

            // Act
            header.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal("Title", cell.StringCellValue);
            var font = workbook.GetFontAt(cell.CellStyle.FontIndex);
            Assert.True(font.IsBold);
        }

        [Fact]
        public void Header_SetBackground_ShouldApplyColor()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var header = new Header("Colored Header").SetBackground(IndexedColors.Blue);

            // Act
            header.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.Blue.Index, cell.CellStyle.FillForegroundColor);
        }

        #endregion

        #region InputCell Widget Tests

        [Fact]
        public void InputCell_Build_ShouldSetDefaultValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var input = new InputCell(defaultValue: 1.5);

            // Act
            input.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(1.5, cell.NumericCellValue);
        }

        [Fact]
        public void InputCell_ShowEditableIndicator_ShouldApplyYellowBackground()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var input = new InputCell(defaultValue: 100);

            // Act
            input.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal(IndexedColors.LightYellow.Index, cell.CellStyle.FillForegroundColor);
        }

        [Fact]
        public void InputCell_WithKey_ShouldAddComment()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var input = new InputCell(defaultValue: 42);
            input.Key = "my_input";

            // Act
            input.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var cell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.NotNull(cell.CellComment);
            Assert.Contains("Key: my_input", cell.CellComment.String.String);
        }

        #endregion

        #region Column Widget Tests

        [Fact]
        public void Column_Build_ShouldArrangeChildrenVertically()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var column = new Column(
                new Label("Row 1"),
                new Label("Row 2"),
                new Label("Row 3")
            );

            // Act
            column.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Row 1", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Row 2", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Row 3", npoiSheet.GetRow(2)?.GetCell(0)?.StringCellValue);
        }

        #endregion

        #region Row Widget Tests

        [Fact]
        public void Row_Build_ShouldArrangeChildrenHorizontally()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);
            var row = new Row(
                new Label("Col A"),
                new Label("Col B"),
                new Label("Col C")
            );

            // Act
            row.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            var npoiRow = npoiSheet.GetRow(0);
            Assert.Equal("Col A", npoiRow?.GetCell(0)?.StringCellValue);
            Assert.Equal("Col B", npoiRow?.GetCell(1)?.StringCellValue);
            Assert.Equal("Col C", npoiRow?.GetCell(2)?.StringCellValue);
        }

        #endregion

        #region Nested Widget Tests

        [Fact]
        public void NestedWidgets_ShouldRenderCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Create a simple table layout
            var table = new Column(
                new Header("Name"),
                new Label("Alice"),
                new Label("Bob")
            );

            // Act
            table.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal("Name", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Alice", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal("Bob", npoiSheet.GetRow(2)?.GetCell(0)?.StringCellValue);
        }

        [Fact]
        public void ComplexLayout_WithColumnsAndRows_ShouldRenderCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");
            var ctx = new ExcelContext(sheet);

            // Create a table header row + data rows
            var layout = new Column(
                new Row(
                    new Header("產品"),
                    new Header("價格"),
                    new Header("數量")
                ),
                new Row(
                    new Label("蘋果"),
                    new InputCell(30),
                    new InputCell(10)
                ),
                new Row(
                    new Label("香蕉"),
                    new InputCell(20),
                    new InputCell(15)
                )
            );

            // Act
            layout.Build(ctx);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");

            // Header row
            Assert.Equal("產品", npoiSheet.GetRow(0)?.GetCell(0)?.StringCellValue);
            Assert.Equal("價格", npoiSheet.GetRow(0)?.GetCell(1)?.StringCellValue);
            Assert.Equal("數量", npoiSheet.GetRow(0)?.GetCell(2)?.StringCellValue);

            // Data row 1
            Assert.Equal("蘋果", npoiSheet.GetRow(1)?.GetCell(0)?.StringCellValue);
            Assert.Equal(30.0, npoiSheet.GetRow(1)?.GetCell(1)?.NumericCellValue);
            Assert.Equal(10.0, npoiSheet.GetRow(1)?.GetCell(2)?.NumericCellValue);

            // Data row 2
            Assert.Equal("香蕉", npoiSheet.GetRow(2)?.GetCell(0)?.StringCellValue);
            Assert.Equal(20.0, npoiSheet.GetRow(2)?.GetCell(1)?.NumericCellValue);
            Assert.Equal(15.0, npoiSheet.GetRow(2)?.GetCell(2)?.NumericCellValue);
        }

        #endregion

        #region ExcelWidget Base Tests

        [Fact]
        public void ExcelWidget_ShouldCaptureSourceLocation()
        {
            // Arrange & Act
            var label = new Label("Test");

            // Assert
            Assert.NotNull(label.SourceLocation);
            Assert.Contains("WidgetRenderingTests.cs", label.SourceLocation.FilePath);
            Assert.True(label.SourceLocation.LineNumber > 0);
        }

        [Fact]
        public void ExcelWidget_WithKey_ShouldStoreKey()
        {
            // Arrange
            var label = new Label("Test");

            // Act
            label.WithKey("my_key");

            // Assert
            Assert.Equal("my_key", label.Key);
        }

        #endregion
    }
}
