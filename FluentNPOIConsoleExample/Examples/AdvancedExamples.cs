using System;
using System.IO;
using System.Collections.Generic;
using FluentNPOI;
using NPOI.XSSF.UserModel;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Advanced Examples - Pipeline, DOM Edit
    /// </summary>
    internal partial class Program
    {
        #region Pipeline and DOM Examples

        /// <summary>
        /// Example 10: Smart Pipeline (Streaming and Legacy)
        /// </summary>
        public static void CreateSmartPipelineExample(List<ExampleData> testData)
        {
            Console.WriteLine("Creating SmartPipelineExample...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A).WithTitle("姓名");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B).WithTitle("分數");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.C).WithTitle("狀態");

            // 1. Generate Source File (Simulation)
            var sourceFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Source.xlsx";
            var wb = new FluentWorkbook(new XSSFWorkbook());
            wb.UseSheet("Data").SetTable(testData, mapping).BuildRows();
            wb.SaveToFile(sourceFile);

            // 2. Stream Processing: Output as .xlsx (SXSSF - High Speed)
            var outFileXlsx = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Pipeline_Out.xlsx";

            FluentNPOI.Streaming.StreamingBuilder<ExampleData>.FromFile(sourceFile)
                .Transform(d =>
                {
                    d.Name += " (Streamed)";
                    d.Score += 1.1;
                })
                .WithMapping(mapping)
                .Configure(sheet =>
                {
                    sheet.SetColumnWidth(ExcelCol.A, 40);
                    sheet.FreezeTitleRow();
                })
                .SaveAs(outFileXlsx);

            Console.WriteLine($"  ✓ Pipeline (XLSX) Processed: {outFileXlsx}");

            // 3. Compatibility Processing: Output as .xls (HSSF - DOM)
            var outFileXls = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Pipeline_Out.xls";

            FluentNPOI.Streaming.StreamingBuilder<ExampleData>.FromFile(sourceFile)
                .Transform(d => d.Name += " (Legacy)")
                .WithMapping(mapping)
                .SaveAs(outFileXls);

            Console.WriteLine($"  ✓ Pipeline (XLS) Processed: {outFileXls}");
        }

        /// <summary>
        /// Example 11: DOM Edit (Modify existing file)
        /// </summary>
        public static void CreateDomEditExample()
        {
            Console.WriteLine("Creating DomEditExample (In-place Edit)...");

            var templateFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Template.xlsx";
            var editedFile = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Edited.xlsx";

            // 1. Prepare a template file
            var wb = new FluentWorkbook(new XSSFWorkbook());
            wb.UseSheet("Report")
              .SetCellPosition(ExcelCol.A, 1).SetValue("Title: Monthly Report")
              .SetCellPosition(ExcelCol.A, 2).SetValue("Generated: [DATE]")
              .SetCellPosition(ExcelCol.B, 5).SetValue("Data Area");

            wb.SaveToFile(templateFile).Close();

            // 2. Load and Edit (Load -> Edit -> Save)
            var editor = new FluentWorkbook(new XSSFWorkbook());
            editor.ReadExcelFile(templateFile);

            editor.UseSheet("Report")
                  .SetCellPosition(ExcelCol.A, 1).SetValue("Title: Final Report 2024")
                  .SetCellPosition(ExcelCol.A, 2).SetValue($"Generated: {DateTime.Now:yyyy-MM-dd}")
                  .SetCellPosition(ExcelCol.A, 10).SetValue("Approved by Manager");

            editor.SaveToFile(editedFile);
            editor.Close();

            Console.WriteLine($"  ✓ DOM Edit Completed: {editedFile}");
        }

        #endregion
    }
}
