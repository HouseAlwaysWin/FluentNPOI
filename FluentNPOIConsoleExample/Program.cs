using System.Data;
using System;
using System.IO;
using FluentNPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Models;
using System.Linq;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Charts;
using FluentNPOI.HotReload;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// FluentNPOI Console Example - Main entry point
    /// Examples are split into partial classes by category
    /// </summary>
    internal partial class Program
    {
        public static List<ExampleData> testData = GetTestData();
        public static string filePath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test.xlsx";
        public static string outputPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\Test2_v2.xlsx";
        static void Main(string[] args)
        {
            // Hot Reload examples - use command line args to select mode
            // dotnet watch run -- widget     # Widget 系統範例
            // dotnet watch run -- fluent     # FluentNPOI 原生範例
            // dotnet watch run -- hybrid     # 混合模式範例
            // dotnet run                     # 一般範例 (無熱重載)
            // if (args.Length > 0)
            // {
            //     switch (args[0].ToLower())
            //     {
            //         case "widget":
            //             HotReloadExamples.RunWidgetExample();
            //             return;
            //         case "fluent":
            //             HotReloadExamples.RunFluentExample();
            //             return;
            //         case "hybrid":
            //             HotReloadExamples.RunHybridExample();
            //             return;
            //         case "style":
            //             HotReloadExamples.RunStyleExample();
            //             return;
            //         case "hotreload":
            //         case "help":
            //             PrintHotReloadHelp();
            //             return;
            //     }
            // }

            // Standard examples (no hot reload)
            try
            {


                // Ensure Resources folder exists
                var dir = Path.GetDirectoryName(outputPath);
                if (dir != null) Directory.CreateDirectory(dir);

                var workbook = new XSSFWorkbook();
                FluentLivePreview.Run(outputPath, wb =>
                {
                    var fluent = new FluentWorkbook(workbook);

                    // Setup styles
                    SetupStyles(fluent);

                    // ========== Write Examples ==========
                    // Table write examples
                    CreateBasicTableExample(fluent, testData);
                    CreateSummaryExample(fluent, testData);
                    CreateDataTableExample(fluent);

                    // Style write examples
                    CreateCopyStyleExample(fluent, testData);
                    CreateCellStyleRangeExample(fluent);
                    CreateSheetGlobalStyleExample(fluent, testData);
                    CreateMappingStylingExample(fluent, testData);

                    // Cell write examples
                    CreateSetCellValueExample(fluent);
                    CreateCellMergeExample(fluent);
                    CreatePictureExample(fluent);

                    // Smart Pipeline example
                    CreateSmartPipelineExample(testData);

                    // DOM Edit example
                    CreateDomEditExample();

                    // HTML Export example
                    CreateHtmlExportExample(fluent);

                    // PDF Export example
                    CreatePdfExportExample(fluent);

                    // Chart example
                    CreateChartExample(fluent, testData);

                    // Save file
                    fluent.SaveToPath(outputPath);
                    Console.WriteLine($"✓ 檔案儲存至: {outputPath}");

                    // Read examples
                    ReadExcelExamples(fluent);
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static void PrintHotReloadHelp()
        {
            Console.WriteLine("FluentNPOI Hot Reload Examples");
            Console.WriteLine("==============================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  dotnet watch run -- widget    Widget 系統範例 (宣告式)");
            Console.WriteLine("  dotnet watch run -- fluent    FluentNPOI 原生範例 (零學習成本)");
            Console.WriteLine("  dotnet watch run -- hybrid    混合模式範例");
            Console.WriteLine("  dotnet run                    一般範例 (無熱重載)");
            Console.WriteLine();
            Console.WriteLine("熱重載使用方式:");
            Console.WriteLine("  1. 執行上述 dotnet watch run 指令");
            Console.WriteLine("  2. 開啟產生的 .xlsx 檔案");
            Console.WriteLine("  3. 修改程式碼並儲存");
            Console.WriteLine("  4. Excel 檔案會自動更新!");
        }

        #region Test Data

        static List<ExampleData> GetTestData()
        {
            return new List<ExampleData>
            {
                new(1, "Alice Chen", new DateTime(1994, 1, 1), true, 95.5, 12500.75m, "優秀學生"),
                new(2, "Bob Lee", new DateTime(1989, 5, 12), false, 78.0, 8900.50m, "需改進"),
                new(3, "Søren", new DateTime(1985, 7, 23), true, 88.5, 15000.00m, "表現良好"),
                new(4, "王小明", new DateTime(2000, 2, 29), true, 92.0, 11200.80m, "進步快速"),
                new(5, "This is a very very very long name to test wrapping and width", new DateTime(1980, 12, 31), false, 65.5, 7500.25m, "LongName"),
                new(6, "Élodie", new DateTime(1995, 5, 15), true, 85.0, 9800.00m, "穩定發揮"),
                new(7, "O'Connor", new DateTime(1975, 7, 7), false, 72.5, 8200.50m, "待觀察"),
                new(8, "李雷", new DateTime(2010, 10, 10), true, 90.0, 10500.75m, "潛力股"),
                new(9, "山田太郎", new DateTime(1999, 3, 3), true, 87.5, 9500.00m, "穩健型"),
                new(10, "Мария", new DateTime(1988, 8, 8), false, 70.0, 8000.25m, "需加強"),
                new(11, "محمد", new DateTime(1991, 9, 9), true, 93.5, 12000.00m, "頂尖"),
                new(12, "김민준", new DateTime(2004, 4, 4), true, 89.0, 10200.50m, "均衡發展"),
            };
        }

        #endregion

        #region Style Setup

        public static void SetupStyles(FluentWorkbook fluent)
        {
            fluent
                // 設定全域基礎樣式
                .SetupGlobalCachedCellStyles((workbook, style) =>
                {
                    style.SetAlignment(HorizontalAlignment.Center);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetFontInfo(workbook, "Calibri", 10);
                })

                // 使用 inheritFrom 繼承 global，只覆寫需要改的屬性
                .SetupCellStyle("BodyString", (workbook, style) =>
                {
                    style.SetFontInfo(workbook, "新細明體", 10);
                }, inheritFrom: "global")

                .SetupCellStyle("DateOfBirth", (workbook, style) =>
                {
                    style.SetDataFormat(workbook, "yyyy-MM-dd");
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                }, inheritFrom: "global")

                .SetupCellStyle("HeaderBlue", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
                }, inheritFrom: "global")

                .SetupCellStyle("BodyGreen", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                }, inheritFrom: "global")

                .SetupCellStyle("AmountCurrency", (workbook, style) =>
                {
                    style.SetDataFormat(workbook, "#,##0.00");
                    style.SetAlignment(HorizontalAlignment.Right);
                }, inheritFrom: "global")

                .SetupCellStyle("HighlightYellow", (workbook, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Yellow);
                }, inheritFrom: "global");
        }

        #endregion
    }
}
