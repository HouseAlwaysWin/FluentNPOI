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
            Console.WriteLine("Creating CellStyleRangeDemo...");

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

            Console.WriteLine("  ✓ CellStyleRangeDemo Created");
        }

        /// <summary>
        /// Example 3.5: Copy style and dynamic style
        /// </summary>
        public static void CreateCopyStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("Creating CopyStyleExample...");

            // Define some dynamic styles
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

            // 1. Copy header style from Sheet1 A1 (HeaderBlue)
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A)
                .WithTitle("Name (Copy Header)")
                .WithRowOffset(2)
                .WithTitleStyleFrom(1, ExcelCol.B);

            // Adjustment: Demonstrate dynamic style directly using style Key
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.B)
                .WithTitle("Status (Dynamic)")
                .WithTitleStyle("HeaderBlue")
                .WithDynamicStyle(item => item.IsActive ? "ActiveGreen" : "InactiveRed");

            // Demonstrate CopyStyleFromCell: Create a template cell in the current Sheet first
            var sheet = fluent.UseSheet("CopyStyleExample", true);
            sheet.SetCellPosition(ExcelCol.Z, 1).SetValue("Template").SetCellStyle("HeaderBlue");

            mapping.Map(x => x.Score).ToColumn(ExcelCol.C)
                .WithTitle("Score (Copy Z1)")
                .WithTitleStyleFrom(1, ExcelCol.Z)
                .WithCellType(CellType.Numeric);

            // Use WithStartRow to set default starting row
            mapping.WithStartRow(2);

            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ CopyStyleExample Created");
        }

        /// <summary>
        /// Example 5: Per-sheet global styles
        /// </summary>
        public static void CreateSheetGlobalStyleExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("Creating SheetGlobalStyleExample...");

            var limitedData = testData.Take(5).ToList();

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID");
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.C).WithTitle("Score")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("Is Active")
                .WithCellType(CellType.Boolean);

            // Sheet 1: Green
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
                .SetValue("Sheet Global Style: Green")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Green")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 2: Yellow
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
                .SetValue("Sheet Global Style: Yellow")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Yellow")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            // Sheet 3: Mix Sheet global style and specific style
            var mixedMapping = new FluentMapping<ExampleData>();
            mixedMapping.Map(x => x.ID).ToColumn(ExcelCol.A).WithTitle("ID")
                .WithStyle("HighlightYellow");

            mixedMapping.Map(x => x.Name).ToColumn(ExcelCol.B).WithTitle("Name");

            mixedMapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C).WithTitle("Date of Birth")
                .WithStyle("DateOfBirth");

            mixedMapping.Map(x => x.IsActive).ToColumn(ExcelCol.D).WithTitle("Is Active")
                .WithCellType(CellType.Boolean)
                .WithDynamicStyle(x => x.IsActive ? "Sheet3_ActiveGreen" : "Sheet3_InactiveRed");

            // Pre-register Dynamic Styles
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
                .SetValue("Mixed Style: Sheet Global (Aqua) + Specific Style Override")
                .SetCellStyle("HeaderBlue");
            fluent.UseSheet("SheetGlobalStyle_Mixed")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1);

            Console.WriteLine("  ✓ SheetGlobalStyleExample Created");
        }

        /// <summary>
        /// Example 9: Direct styling in mapping (New Feature)
        /// </summary>
        public static void CreateMappingStylingExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("Creating MappingStylingExample...");

            var mapping = new FluentMapping<ExampleData>();

            mapping.Map(x => x.Name)
                .ToColumn(ExcelCol.A)
                .WithTitle("Name")
                .WithBackgroundColor(IndexedColors.LightCornflowerBlue)
                .WithAlignment(HorizontalAlignment.Center);

            mapping.Map(x => x.Score)
                .ToColumn(ExcelCol.B)
                .WithTitle("Score")
                .WithNumberFormat("0.0")
                .WithFont(isBold: true);

            mapping.Map(x => x.DateOfBirth)
                .ToColumn(ExcelCol.C)
                .WithTitle("Date")
                .WithNumberFormat("yyyy-mm-dd")
                .WithBackgroundColor(IndexedColors.LightYellow);

            fluent.UseSheet("MappingStylingExample", true)
               .SetColumnWidth(ExcelCol.A, ExcelCol.C, 20)
               .SetTable(testData.Take(5), mapping)
               .BuildRows()
               .SetAutoFilter();

            Console.WriteLine("  ✓ MappingStylingExample Created");
        }

        /// <summary>
        /// Example 10: DataTable Styling (New Feature)
        /// </summary>
        public static void CreateDataTableStyleExample(FluentWorkbook fluent)
        {
            Console.WriteLine("Creating DataTableStyleExample...");

            var dt = new System.Data.DataTable();
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Price", typeof(decimal));
            dt.Columns.Add("InStock", typeof(bool));

            dt.Rows.Add("Laptop", 1200.00m, true);
            dt.Rows.Add("Mouse", 25.50m, true);
            dt.Rows.Add("Keyboard", 45.00m, false);

            var mapping = DataTableMapping.FromDataTable(dt);

            // Customize mapping with styles
            mapping.Map("Product")
                .ToColumn(ExcelCol.A)
                .WithTitle("Product Name")
                .WithTitleFont(fontName: "Arial", isBold: true, fontSize: 14)
                .WithTitleBackgroundColor(IndexedColors.Grey25Percent)
                .WithFont(isBold: true, fontSize: 12)
                .WithBackgroundColor(IndexedColors.LightCornflowerBlue)
                .WithAlignment(HorizontalAlignment.Center);

            mapping.Map("Price")
                .ToColumn(ExcelCol.B)
                .WithTitle("Price (USD)")
                .WithTitleFont(color: IndexedColors.White)
                .WithTitleBackgroundColor(IndexedColors.DarkBlue)
                .WithTitleAlignment(HorizontalAlignment.Right)
                .WithNumberFormat("$#,##0.00")
                .WithFont(color: IndexedColors.Green);

            mapping.Map("InStock")
                .ToColumn(ExcelCol.C)
                .WithTitle("Available")
                .WithAlignment(HorizontalAlignment.Center)
                .WithDynamicStyle(row => 
                {
                    // Dynamic style based on value (DataRow)
                    var inStock = (bool)row["InStock"];
                    return inStock ? "InStockStyle" : "OutOfStockStyle";
                });

            // Register dynamic styles
            fluent.SetupCellStyle("InStockStyle", (wb, s) => 
            {
                s.SetFontInfo(wb, color: IndexedColors.Green);
            });
            fluent.SetupCellStyle("OutOfStockStyle", (wb, s) =>
            {
                s.SetFontInfo(wb, color: IndexedColors.Red, isBold: true);
            });

            fluent.UseSheet("DataTableStyleExample", true)
                .WriteDataTable(dt, mapping)
                .SetColumnWidth(ExcelCol.A, ExcelCol.C, 15);

            Console.WriteLine("  ✓ DataTableStyleExample Created");
        }

        #endregion
    }
}
