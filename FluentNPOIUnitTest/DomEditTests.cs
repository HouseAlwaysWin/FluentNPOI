using System;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;
using FluentNPOI;
using FluentNPOI.Stages;

namespace FluentNPOIUnitTest
{
    public class DomEditTests
    {
        [Fact]
        public void ReadExcelFile_ShouldPreserveOriginalData_WhenEditing()
        {
            // Arrange
            var templateFile = "template.xlsx";
            var editedFile = "edited.xlsx";

            // 1. Create a "Template" with specific data
            var wb = new FluentWorkbook(new XSSFWorkbook());
            wb.UseSheet("Sheet1")
              .SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1).SetValue("Original A1")
              .SetCellPosition(FluentNPOI.Models.ExcelCol.B, 1).SetValue("Original B1");
            wb.SaveToFile(templateFile).Close();

            // Act
            // 2. Load using ReadExcelFile (DOM Mode)
            var fluent = new FluentWorkbook(new XSSFWorkbook()); // Dummy init to access method, wait.
                                                                 // FluentWorkbook constructor requires IWorkbook. 
                                                                 // ReadExcelFile is an instance method? Let's check.
                                                                 // Yes, user pointed to public FluentWorkbook ReadExcelFile(string).
                                                                 // But wait, to call ReadExcelFile, I need an instance?
                                                                 // "public FluentWorkbook ReadExcelFile(string filePath)"
                                                                 // If I create 'new FluentWorkbook(null)', it might crash.
                                                                 // Actually, ReadExcelFile is designed to RE-LOAD into current wrapper?
                                                                 // Let's check implementation.
            /*
             public FluentWorkbook ReadExcelFile(string filePath) {
                ...
                return ReadExcelStream(fs);
             }
             public FluentWorkbook ReadExcelStream(Stream stream) {
                _workbook?.Close(); ... _workbook = WorkbookFactory.Create(stream); ...
             }
            */
            // So yes, I can reuse an instance or create a dummy one.
            // Ideally, FluentWorkbook might need a parameterless constructor or static factory?
            // Currently, generic constructor takes IWorkbook. 
            // I'll use `new FluentWorkbook(null)` if allowed, or `new FluentWorkbook(new XSSFWorkbook())`.

            var editor = new FluentWorkbook(new XSSFWorkbook()); // Start with empty
            editor.ReadExcelFile(templateFile)
                  .UseSheet("Sheet1")
                  .SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1)
                  .SetValue("Edited A1"); // Modify A1 only

            editor.SaveToFile(editedFile);
            editor.Close();

            // Assert
            var verifyWb = new XSSFWorkbook(File.OpenRead(editedFile));
            var sheet = verifyWb.GetSheet("Sheet1");

            // A1 should be changed
            Assert.Equal("Edited A1", sheet.GetRow(0).GetCell(0).StringCellValue);

            // B1 should remain "Original B1" (Prove we didn't wipe the sheet)
            Assert.Equal("Original B1", sheet.GetRow(0).GetCell(1).StringCellValue);

            // Cleanup
            if (File.Exists(templateFile)) File.Delete(templateFile);
            if (File.Exists(editedFile)) File.Delete(editedFile);
        }
    }
}
