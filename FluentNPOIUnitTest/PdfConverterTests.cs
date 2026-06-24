using System.IO;
using System.Text;
using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Pdf;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FluentNPOIUnitTest
{
    public class PdfConverterTests
    {
        private static XSSFWorkbook BuildSimpleWorkbook()
        {
            var workbook = new XSSFWorkbook();
            new FluentWorkbook(workbook)
                .UseSheet("Sheet1")
                .SetCellPosition(ExcelCol.A, 1).SetValue("Name")
                .SetCellPosition(ExcelCol.B, 1).SetValue("Age")
                .SetCellPosition(ExcelCol.A, 2).SetValue("Alice")
                .SetCellPosition(ExcelCol.B, 2).SetValue(30);
            return workbook;
        }

        private static bool HasPdfHeader(byte[] bytes)
        {
            return bytes.Length >= 4 && Encoding.ASCII.GetString(bytes, 0, 4) == "%PDF";
        }

        [Fact]
        public void ConvertSheetToPdfBytes_ProducesValidPdf()
        {
            var workbook = BuildSimpleWorkbook();
            var sheet = workbook.GetSheetAt(0);

            var bytes = PdfConverter.ConvertSheetToPdfBytes(sheet, workbook);

            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0, "Expected non-empty PDF output");
            Assert.True(HasPdfHeader(bytes), "Output should start with the %PDF magic header");
        }

        [Fact]
        public void ConvertSheetToPdfBytes_WithMergedCellsAndStyles_DoesNotThrow()
        {
            var workbook = new XSSFWorkbook();
            var sheet = new FluentWorkbook(workbook).UseSheet("Styled");
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Merged Title")
                 .SetFont(isBold: true)
                 .SetBackgroundColor(IndexedColors.Yellow);
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1, 1);
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("Cell");

            var npoiSheet = workbook.GetSheetAt(0);
            var bytes = PdfConverter.ConvertSheetToPdfBytes(npoiSheet, workbook);

            Assert.True(bytes.Length > 0);
            Assert.True(HasPdfHeader(bytes));
        }

        [Fact]
        public void ConvertSheetToPdf_WritesNonEmptyFile()
        {
            var workbook = BuildSimpleWorkbook();
            var sheet = workbook.GetSheetAt(0);
            var path = Path.Combine(Path.GetTempPath(), $"pdf_test_{System.Guid.NewGuid()}.pdf");

            try
            {
                PdfConverter.ConvertSheetToPdf(sheet, workbook, path);

                Assert.True(File.Exists(path), "PDF file should be created");
                Assert.True(new FileInfo(path).Length > 0, "PDF file should be non-empty");
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [Fact]
        public void ConvertSheetToPdfBytes_EmptySheet_DoesNotThrow()
        {
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("Empty");
            var sheet = workbook.GetSheetAt(0);

            // An empty sheet has no columns; conversion should still succeed without throwing.
            var bytes = PdfConverter.ConvertSheetToPdfBytes(sheet, workbook);

            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
        }
    }
}
