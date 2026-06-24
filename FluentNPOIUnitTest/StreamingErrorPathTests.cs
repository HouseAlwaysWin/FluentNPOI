using System;
using System.IO;
using Xunit;
using FluentNPOI.Streaming;
using FluentNPOI.Streaming.Rows;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    public class StreamingErrorPathTests
    {
        public class Rec
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
        }

        private static string CreateXlsx()
        {
            var path = Path.Combine(Path.GetTempPath(), $"sdt_{Guid.NewGuid()}.xlsx");
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Header");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            return path;
        }

        [Fact]
        public void StreamingBuilder_FromFile_NullOrEmpty_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => StreamingBuilder<Rec>.FromFile(null!));
            Assert.Throws<ArgumentNullException>(() => StreamingBuilder<Rec>.FromFile(""));
        }

        [Fact]
        public void StreamingBuilder_FromStream_Null_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => StreamingBuilder<Rec>.FromStream(null!));
        }

        [Fact]
        public void StreamingBuilder_SaveAs_NullOrEmpty_Throws()
        {
            using var ms = new MemoryStream();
            var builder = StreamingBuilder<Rec>.FromStream(ms);
            Assert.Throws<ArgumentNullException>(() => builder.SaveAs(null!));
            Assert.Throws<ArgumentNullException>(() => builder.SaveAs(""));
        }

        [Fact]
        public void StreamingRow_GetValueT_InvalidConversion_ReturnsDefault()
        {
            var row = new StreamingRow(0, new object[] { "not-a-number" });
            // Convert.ChangeType throws internally; GetValue<T> swallows and returns default.
            Assert.Equal(0, row.GetValue<int>(0));
        }

        [Fact]
        public void StreamingRow_GetValueT_OutOfRange_ReturnsDefault()
        {
            var row = new StreamingRow(0, new object[] { "x" });
            Assert.Equal(0, row.GetValue<int>(5));
            Assert.Null(row.GetValue<string>(5));
        }

        [Fact]
        public void ReadAsDataTable_MissingSheet_ReturnsNull()
        {
            var path = CreateXlsx();
            try
            {
                var table = FluentExcelReader.ReadAsDataTable(path, "NoSuchSheet");
                Assert.Null(table);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAsDataTable_ExistingSheet_ReturnsTable()
        {
            var path = CreateXlsx();
            try
            {
                var table = FluentExcelReader.ReadAsDataTable(path, "Data");
                Assert.NotNull(table);
                Assert.True(table!.Rows.Count >= 1);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
