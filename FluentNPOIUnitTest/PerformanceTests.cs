using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using FluentNPOI.Streaming;
using NPOI.XSSF.UserModel;
using Xunit;
using Xunit.Abstractions;

namespace FluentNPOIUnitTest
{
    public class PerformanceTests
    {
        private readonly ITestOutputHelper _output;

        public PerformanceTests(ITestOutputHelper output)
        {
            _output = output;
        }

        public class BenchModel
        {
            public int Id { get; set; }
            public string? Name { get; set; }
            public string? Email { get; set; }
            public double Score { get; set; }
            public DateTime CreatedAt { get; set; }
        }

        [Fact]
        public void Read_100k_Rows_Benchmark()
        {
            int rowCount = 100_000;
            string filePath = Path.Combine(Path.GetTempPath(), $"bench_{Guid.NewGuid()}.xlsx");

            try
            {
                // 1. Setup: Generate large Excel file
                _output.WriteLine($"Generating {rowCount} rows...");
                var swSetup = Stopwatch.StartNew();
                CreateLargeExcelFile(filePath, rowCount);
                swSetup.Stop();
                _output.WriteLine($"Generation took: {swSetup.ElapsedMilliseconds} ms");

                // 2. Measure Read
                _output.WriteLine("Starting Read benchmark...");
                var swRead = Stopwatch.StartNew();

                // Force enumeration to ensure all rows are read
                var results = FluentExcelReader.Read<BenchModel>(filePath).ToList();

                swRead.Stop();

                // 3. Report
                double seconds = swRead.Elapsed.TotalSeconds;
                double rowsPerSec = rowCount / seconds;

                _output.WriteLine($"--------------------------------------------------");
                _output.WriteLine($"Read {results.Count} rows in {seconds:F4} seconds");
                _output.WriteLine($"Throughput: {rowsPerSec:N0} rows/sec");
                _output.WriteLine($"--------------------------------------------------");

                Assert.Equal(rowCount, results.Count);
                Assert.Equal(0, results[0].Id);
                Assert.Equal("User 0", results[0].Name);
                Assert.Equal(rowCount - 1, results[rowCount - 1].Id);
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    try { File.Delete(filePath); } catch { }
                }
            }
        }

        private void CreateLargeExcelFile(string filePath, int rowCount)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");

            // Header
            var headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Id");
            headerRow.CreateCell(1).SetCellValue("Name");
            headerRow.CreateCell(2).SetCellValue("Email");
            headerRow.CreateCell(3).SetCellValue("Score");
            headerRow.CreateCell(4).SetCellValue("CreatedAt");

            // Data
            // Note: NPOI XSSF can be slow for writing 100k rows (DOM based). 
            // In a real scenario we might want to use SXSSF (Streamed) for writing, 
            // but here we just need a file. We'll stick to simple XSSF but be patient on setup.
            var now = DateTime.Now;
            for (int i = 0; i < rowCount; i++)
            {
                var row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(i);
                row.CreateCell(1).SetCellValue($"User {i}");
                row.CreateCell(2).SetCellValue($"user{i}@example.com");
                row.CreateCell(3).SetCellValue(i * 0.5);
                row.CreateCell(4).SetCellValue(now.AddSeconds(i).ToString("yyyy-MM-dd HH:mm:ss"));
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
}
