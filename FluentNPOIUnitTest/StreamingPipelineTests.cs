using System;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using Xunit;
using FluentNPOI;
using FluentNPOI.Stages;
using FluentNPOI.Streaming;

namespace FluentNPOIUnitTest
{
    public class StreamingPipelineTests
    {
        private class TestData
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public double Value { get; set; }
        }

        [Fact]
        public void Process_ShouldTransformAndWriteData_WhenPipelineIsRun()
        {
            // Arrange
            var inputFile = "pipeline_in.xlsx";
            var outputFile = "pipeline_out.xlsx";

            // Create input file with regular FluentWorkbook (ok for small test data)
            var inputData = Enumerable.Range(1, 10).Select(i => new TestData
            {
                Id = i,
                Name = $"Item {i}",
                Value = i * 10
            });

            var wb = new FluentWorkbook(new XSSFWorkbook());

            // Define mapping
            var mapping = new FluentNPOI.Streaming.Mapping.FluentMapping<TestData>();
            mapping.Map(x => x.Id).ToColumn(FluentNPOI.Models.ExcelCol.A).WithTitle("Id");
            mapping.Map(x => x.Name).ToColumn(FluentNPOI.Models.ExcelCol.B).WithTitle("Name");
            mapping.Map(x => x.Value).ToColumn(FluentNPOI.Models.ExcelCol.C).WithTitle("Value");

            wb.UseSheet("Sheet1")
              .SetTable(inputData, mapping)
              .BuildRows();

            wb.SaveToFile(inputFile)
              .Close();

            // Act
            // Run Pipeline: Read -> Double the Value -> Write
            FluentWorkbook.Stream<TestData>(inputFile)
                .Transform((item) =>
                {
                    item.Value = item.Value * 2;
                    item.Name = item.Name + " (Processed)";
                })
                .WithMapping(mapping) // Use the same mapping
                .Configure(sheet =>
                {
                    // Verify we can use FluentSheet API
                    sheet.SetColumnWidth(FluentNPOI.Models.ExcelCol.A, 50);
                })
                .SaveAs(outputFile);

            // Assert
            // Read back to verify using FluentExcelReader directly
            var result = FluentExcelReader.Read<TestData>(outputFile).ToList();

            Assert.Equal(10, result.Count);
            Assert.Equal(20, result[0].Value); // 10 * 2
            Assert.Equal("Item 1 (Processed)", result[0].Name);
            Assert.Equal(200, result[9].Value); // 100 * 2

            // Cleanup
            if (File.Exists(outputFile)) File.Delete(outputFile);
        }

        [Fact]
        public void Process_ShouldWorkForLegacyXls_WhenPipelineIsRun()
        {
            // Arrange
            var inputFile = "legacy_in.xls";   // .xls extension
            var outputFile = "legacy_out.xls"; // .xls extension (triggers DOM backend)

            var inputData = Enumerable.Range(1, 5).Select(i => new TestData
            {
                Id = i,
                Name = $"Old {i}",
                Value = i * 100
            });

            // Create .xls file using HSSFWorkbook explicitly
            var wb = new FluentWorkbook(new NPOI.HSSF.UserModel.HSSFWorkbook());

            var mapping = new FluentNPOI.Streaming.Mapping.FluentMapping<TestData>();
            mapping.Map(x => x.Id).ToColumn(FluentNPOI.Models.ExcelCol.A).WithTitle("Id");
            mapping.Map(x => x.Name).ToColumn(FluentNPOI.Models.ExcelCol.B).WithTitle("Name");
            mapping.Map(x => x.Value).ToColumn(FluentNPOI.Models.ExcelCol.C).WithTitle("Value");

            wb.UseSheet("Sheet1")
              .SetTable(inputData, mapping)
              .BuildRows();

            wb.SaveToFile(inputFile)
              .Close();

            // Act
            // Use the same Stream API
            FluentWorkbook.Stream<TestData>(inputFile)
                .Transform(x => x.Name += " (Updated)")
                .WithMapping(mapping)
                .SaveAs(outputFile);

            // Assert
            var result = FluentExcelReader.Read<TestData>(outputFile).ToList();
            Assert.Equal(5, result.Count);
            Assert.Equal("Old 1 (Updated)", result[0].Name);

            if (File.Exists(inputFile)) File.Delete(inputFile);
            if (File.Exists(outputFile)) File.Delete(outputFile);
        }
    }
}
