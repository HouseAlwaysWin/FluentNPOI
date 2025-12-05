using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using System;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class PositionNormalizationTests
    {
        [Fact]
        public void SetCellPosition_WithOneBased_ShouldConvertToZeroBased()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 使用 1-based 行號
            fluentWorkbook.UseSheet("Sheet1")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Row 1");

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
            Assert.Equal("Row 1", cell.StringCellValue);
        }

        [Fact]
        public void SetCellPosition_WithNegativeRow_ShouldUseZero()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetCellPosition(ExcelCol.A, -1)
                .SetValue("Test");

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var cell = sheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(cell);
        }
        [Fact]
        public void GetCellValue_ShouldReturnCorrectValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // 為日期單元格設置格式
            fluentWorkbook.SetupCellStyle("DateFormat", (wb, style) =>
            {
                style.SetDataFormat(wb, "yyyy-MM-dd");
            });

            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // 設置不同類型的值
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Hello");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(45.67);
            sheet.SetCellPosition(ExcelCol.D, 1).SetValue(true);

            // 設置日期值並套用日期格式
            sheet.SetCellPosition(ExcelCol.E, 1)
                .SetValue(new DateTime(2024, 1, 15))
                .SetCellStyle("DateFormat");

            // Act & Assert
            var stringValue = sheet.GetCellValue<string>(ExcelCol.A, 1);
            Assert.Equal("Hello", stringValue);

            var intValue = sheet.GetCellValue<int>(ExcelCol.B, 1);
            Assert.Equal(123, intValue);

            var doubleValue = sheet.GetCellValue<double>(ExcelCol.C, 1);
            Assert.Equal(45.67, doubleValue, 2);

            var boolValue = sheet.GetCellValue<bool>(ExcelCol.D, 1);
            Assert.True(boolValue);

            var dateValue = sheet.GetCellValue<DateTime>(ExcelCol.E, 1);
            Assert.Equal(new DateTime(2024, 1, 15), dateValue);
        }

        [Fact]
        public void GetCellValue_NonExistentCell_ShouldReturnDefault()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // Act
            var value = sheet.GetCellValue<string>(ExcelCol.Z, 100);

            // Assert
            Assert.Null(value);
        }

        [Fact]
        public void FluentCell_GetValue_ShouldReturnCorrectValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // Act
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Test Value");
            var cell = sheet.SetCellPosition(ExcelCol.A, 1);
            var value = cell.GetValue<string>();

            // Assert
            Assert.Equal("Test Value", value);
        }

        [Fact]
        public void SetAndGetFormula_ShouldWork()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            // 設置一些數值
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue(10);
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue(20);

            // Act - 設置公式
            sheet.SetCellPosition(ExcelCol.C, 1).SetFormulaValue("=A1+B1");

            // 讀取公式
            var formula = sheet.GetCellFormula(ExcelCol.C, 1);

            // Assert
            Assert.Equal("A1+B1", formula);
        }

        [Fact]
        public void GetCellValue_WithObject_ShouldReturnCorrectType()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("TestSheet");

            sheet.SetCellPosition(ExcelCol.A, 1).SetValue(123);
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue("Text");
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue(true);

            // Act
            var numValue = sheet.GetCellValue(ExcelCol.A, 1);
            var textValue = sheet.GetCellValue(ExcelCol.B, 1);
            var boolValue = sheet.GetCellValue(ExcelCol.C, 1);

            // Assert
            Assert.IsType<double>(numValue);
            Assert.IsType<string>(textValue);
            Assert.IsType<bool>(boolValue);
        }
    }
}

