using System;
using System.IO;
using System.Data;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace FluentNPOIUnitTest
{
    public class DataTableStyleTests
    {
        [Fact]
        public void DataTable_WithMapping_WithFont_StylesApplied()
        {
            var testFilePath = Path.Combine(Path.GetTempPath(), $"dt_style_font_{Guid.NewGuid()}.xlsx");
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Rows.Add("Test Name");

                var mapping = new DataTableMapping();
                mapping.Map("Name")
                    .ToColumn(ExcelCol.A)
                    .WithTitle("Name")
                    .WithFont(fontSize: 14, isBold: true);

                var workbook = new XSSFWorkbook();
                new FluentWorkbook(workbook)
                    .UseSheet("Sheet1")
                    .WriteDataTable(dt, mapping);

                using (var fs = new FileStream(testFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                using (var fs = new FileStream(testFilePath, FileMode.Open, FileAccess.Read))
                {
                    var readWorkbook = new XSSFWorkbook(fs);
                    var sheet = readWorkbook.GetSheetAt(0);
                    var cell = sheet.GetRow(1).GetCell(0);
                    var font = cell.CellStyle.GetFont(readWorkbook);

                    Assert.Equal(14, font.FontHeightInPoints);
                    Assert.True(font.IsBold);
                }
            }
            finally
            {
                if (File.Exists(testFilePath))
                    File.Delete(testFilePath);
            }
        }

        [Fact]
        public void DataTable_WithMapping_WithBackgroundColor_StylesApplied()
        {
            var testFilePath = Path.Combine(Path.GetTempPath(), $"dt_style_bg_{Guid.NewGuid()}.xlsx");
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("Status", typeof(string));
                dt.Rows.Add("Active");

                var mapping = new DataTableMapping();
                mapping.Map("Status")
                    .ToColumn(ExcelCol.B)
                    .WithTitle("Status")
                    .WithBackgroundColor(IndexedColors.Yellow);

                var workbook = new XSSFWorkbook();
                new FluentWorkbook(workbook)
                    .UseSheet("Sheet1")
                    .WriteDataTable(dt, mapping);

                using (var fs = new FileStream(testFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                using (var fs = new FileStream(testFilePath, FileMode.Open, FileAccess.Read))
                {
                    var readWorkbook = new XSSFWorkbook(fs);
                    var sheet = readWorkbook.GetSheetAt(0);
                    var cell = sheet.GetRow(1).GetCell(1);

                    Assert.Equal(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
                    Assert.Equal(IndexedColors.Yellow.Index, cell.CellStyle.FillForegroundColor);
                }
            }
            finally
            {
                if (File.Exists(testFilePath))
                    File.Delete(testFilePath);
            }
        }

        [Fact]
        public void DataTable_WithMapping_CombinedStyles_Works()
        {
            var testFilePath = Path.Combine(Path.GetTempPath(), $"dt_style_combined_{Guid.NewGuid()}.xlsx");
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("Value", typeof(int));
                dt.Rows.Add(12345);

                var mapping = new DataTableMapping();
                mapping.Map("Value")
                    .ToColumn(ExcelCol.A)
                    .WithTitle("Value")
                    .WithFont(isBold: true)
                    .WithAlignment(HorizontalAlignment.Center)
                    .WithNumberFormat("#,##0");

                var workbook = new XSSFWorkbook();
                new FluentWorkbook(workbook)
                    .UseSheet("Sheet1")
                    .WriteDataTable(dt, mapping);

                using (var fs = new FileStream(testFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                using (var fs = new FileStream(testFilePath, FileMode.Open, FileAccess.Read))
                {
                    var readWorkbook = new XSSFWorkbook(fs);
                    var sheet = readWorkbook.GetSheetAt(0);
                    var cell = sheet.GetRow(1).GetCell(0);
                    var font = cell.CellStyle.GetFont(readWorkbook);

                    Assert.True(font.IsBold);
                    Assert.Equal(HorizontalAlignment.Center, cell.CellStyle.Alignment);
                    // DataFormat check might differ based on built-in formats, but checking value is formatted when opened in Excel is the goal.
                    // For automated test, we verify style DataFormat index is set.
                    Assert.True(cell.CellStyle.DataFormat > 0); 
                }
            }
            finally
            {
                if (File.Exists(testFilePath))
                    File.Delete(testFilePath);
            }
        }
    }
}
