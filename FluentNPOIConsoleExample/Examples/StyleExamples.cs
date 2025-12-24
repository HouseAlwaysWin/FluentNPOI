using System;
using FluentNPOI;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Models;
using System.Linq;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Style Write Examples - Cell styling, dynamic styles, sheet global styles
    /// </summary>
    internal partial class Program
    {
        #region Style Write Examples

        /// <summary>
        /// Example 4: Batch set cell range styles
        /// </summary>
        public static void CreateCellStyleRangeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 CellStyleRangeDemo...");

            fluent.UseSheet("CellStyleRangeDemo", true)
                .SetCellStyleRange(new CellStyleConfig("HighlightRed", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Red);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 1, 3)
                .SetCellStyleRange(new CellStyleConfig("HighlightOrange", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Orange);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 4, 6)
                .SetCellStyleRange(new CellStyleConfig("HighlightYellow", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Yellow);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 7, 9)
                .SetCellStyleRange(new CellStyleConfig("HighlightGreen", style =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Green);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                }), ExcelCol.A, ExcelCol.D, 10, 12);

            Console.WriteLine("  ✓ CellStyleRangeDemo 建立完成");
        }

        /// <summary>
        /// Example 3.5: Copy style and dynamic style
        /// </summary>
        public static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 CopyStyleExample...");

            // 定義幾個動態樣式
            fluent.SetupCellStyle("ActiveGreen", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                style.SetBorderAllStyle(BorderStyle.Thin);
                style.SetAlignment(HorizontalAlignment.Center);
            })
            .SetupCellStyle("InactiveRed", (wb, style) =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.SetCellFillForegroundColor(IndexedColors.Rose);
                style.SetBorderAllStyle(BorderStyle.Thin);
                style.SetAlignment(HorizontalAlignment.Center);
            });

            var mapping = new FluentMapping<ExampleData>();

            // 1. 從 Sheet1 的 A1 複製標題樣式 (HeaderBlue)
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A)
                .WithTitle("姓名 (Copy Header)")
                .WithRowOffset(2)
                .WithTitleStyleFrom(1, ExcelCol.B);

            // 調整: 直接用樣式 Key 演示 dynamic style
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.B)
                .WithTitle("狀態 (Dynamic)")
                .WithTitleStyle("HeaderBlue")
                .WithDynamicStyle(item => item.IsActive ? "ActiveGreen" : "InactiveRed");

            // 演示 CopyStyleFromCell: 先在目前 Sheet 建立一個樣板儲存格
            var sheet = fluent.UseSheet("CopyStyleExample", true);
            sheet.SetCellPosition(ExcelCol.Z, 1).SetValue("Template").SetCellStyle("HeaderBlue");

            mapping.Map(x => x.Score).ToColumn(ExcelCol.C)
                .WithTitle("分數 (Copy Z1)")
                .WithTitleStyleFrom(1, ExcelCol.Z)
                .WithCellType(CellType.Numeric);

            // 使用 WithStartRow 設定預設起始列
            mapping.WithStartRow(2);

            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ CopyStyleExample 建立完成");
        }

        /// <summary>
        /// Example 5: Per-sheet global styles
        /// </summary>
        public static void CreateSheetGlobalStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 SheetGlobalStyle...");

            var limitedData = testData.Take(5).ToList();

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("名稱");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.C).WithTitle("分數")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("是否活躍")
                .WithCellType(CellType.Boolean);

            // Sheet 1: 綠色
            fluent.UseSheet("SheetGlobalStyle_Green", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightGreen);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                })
                .SetTable(limitedData, mapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Green")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Sheet 全域樣式：綠色")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Green")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 2: 黃色
            fluent.UseSheet("SheetGlobalStyle_Yellow", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.LightYellow);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                })
                .SetTable(limitedData, mapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Yellow")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Sheet 全域樣式：黃色")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Yellow")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 3: 混合使用 Sheet 全域樣式和特定樣式
            var mixedMapping = new FluentMapping<ExampleData>();
            mixedMapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID")
                .WithStyle("HighlightYellow");

            mixedMapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("名稱");

            mixedMapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C).WithTitle("生日")
                .WithStyle("DateOfBirth");

            mixedMapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("是否活躍")
                .WithCellType(CellType.Boolean)
                .WithDynamicStyle(x => x.IsActive ? "Sheet3_ActiveGreen" : "Sheet3_InactiveRed");

            // 預先註冊 Dynamic Styles
            fluent.SetupCellStyle("Sheet3_ActiveGreen", (wb, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Green);
                s.SetBorderAllStyle(BorderStyle.Thin);
                s.SetFontInfo(wb, fontFamily: "新細明體");
            });
            fluent.SetupCellStyle("Sheet3_InactiveRed", (wb, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Red);
                s.SetBorderAllStyle(BorderStyle.Thin);
                s.SetFontInfo(wb, fontFamily: "新細明體");
            });

            fluent.UseSheet("SheetGlobalStyle_Mixed", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.D, 20)
                .SetupSheetGlobalCachedCellStyles((wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.SetCellFillForegroundColor(IndexedColors.Aqua);
                    style.SetBorderAllStyle(BorderStyle.Thin);
                    style.SetAlignment(HorizontalAlignment.Center);
                    style.SetFontInfo(wb, fontFamily: "新細明體");
                })
                .SetTable(limitedData, mixedMapping, 2)
                .BuildRows();

            fluent.UseSheet("SheetGlobalStyle_Mixed")
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("混合樣式：Sheet 全域(水藍) + 特定樣式覆蓋")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Mixed")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            Console.WriteLine("  ✓ SheetGlobalStyle 建立完成");
        }

        /// <summary>
        /// Example 9: Direct styling in mapping (New Feature)
        /// </summary>
        public static void CreateMappingStylingExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 MappingStylingExample...");

            var mapping = new FluentMapping<ExampleData>();

            mapping.Map(x => x.Name)
                .ToColumn(ExcelCol.A)
                .WithTitle("姓名")
                .WithBackgroundColor(IndexedColors.LightCornflowerBlue)
                .WithAlignment(HorizontalAlignment.Center);

            mapping.Map(x => x.Score)
                .ToColumn(ExcelCol.B)
                .WithTitle("分數")
                .WithNumberFormat("0.0")
                .WithFont(isBold: true);

            mapping.Map(x => x.DateOfBirth)
                .ToColumn(ExcelCol.C)
                .WithTitle("日期")
                .WithNumberFormat("yyyy-mm-dd")
                .WithBackgroundColor(IndexedColors.LightYellow);

            fluent.UseSheet("MappingStylingExample", true)
               .SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
               .SetTable(testData.Take(5), mapping)
               .BuildRows()
               .SetAutoFilter();

            Console.WriteLine("  ✓ MappingStylingExample 建立完成");
        }

        #endregion
    }
}
