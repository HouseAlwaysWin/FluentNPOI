using System;
using System.IO;
using FluentNPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using FluentNPOI.Models;
using FluentNPOI.Stages;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Cell Write Examples - Single cell, merge, pictures
    /// </summary>
    internal partial class Program
    {
        #region Cell Write Examples

        /// <summary>
        /// Example 6: Set single cell value
        /// </summary>
        static void CreateSetCellValueExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 SetCellValueExample...");

            fluent.UseSheet("SetCellValueExample", true)
                .SetColumnWidth(ExcelCol.A, 20)
                .SetCellPosition(ExcelCol.A, 1)
                .SetValue("Hello, World!")
                .SetCellStyle("HighlightYellow");

            Console.WriteLine("  ✓ SetCellValueExample 建立完成");
        }

        /// <summary>
        /// Example 7: Cell merge
        /// </summary>
        static void CreateCellMergeExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 CellMergeExample...");

            var sheet = fluent.UseSheet("CellMergeExample", true);
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.E, 15);

            // 水平合併
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("銷售報表")
                .SetCellStyle("HeaderBlue");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 1);

            // 設定子標題
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("產品名稱");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("銷售量");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue("單價");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue("總金額");
            sheet.SetCellPosition(ExcelCol.E, 2).SetValue("備註");

            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.E; col++)
            {
                sheet.SetCellPosition(col, 2).SetCellStyle("HeaderBlue");
            }

            // 垂直合併
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("電子產品");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 3, 5);

            sheet.SetCellPosition(ExcelCol.B, 3).SetValue(100);
            sheet.SetCellPosition(ExcelCol.C, 3).SetValue(5000);
            sheet.SetCellPosition(ExcelCol.D, 3).SetValue(500000);
            sheet.SetCellPosition(ExcelCol.E, 3).SetValue("熱銷");

            sheet.SetCellPosition(ExcelCol.B, 4).SetValue(80);
            sheet.SetCellPosition(ExcelCol.C, 4).SetValue(3000);
            sheet.SetCellPosition(ExcelCol.D, 4).SetValue(240000);
            sheet.SetCellPosition(ExcelCol.E, 4).SetValue("穩定");

            sheet.SetCellPosition(ExcelCol.B, 5).SetValue(50);
            sheet.SetCellPosition(ExcelCol.C, 5).SetValue(2000);
            sheet.SetCellPosition(ExcelCol.D, 5).SetValue(100000);
            sheet.SetCellPosition(ExcelCol.E, 5).SetValue("一般");

            // 區域合併
            sheet.SetCellPosition(ExcelCol.A, 6).SetValue("總計");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue(840000);
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 6);
            sheet.SetCellPosition(ExcelCol.A, 6).SetCellStyle("HighlightYellow");

            // Multiple merge regions example
            sheet.SetCellPosition(ExcelCol.A, 8).SetValue("部門A");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 8, 10);

            sheet.SetCellPosition(ExcelCol.B, 8).SetValue("部門B");
            sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.B, 8, 10);

            sheet.SetCellPosition(ExcelCol.C, 8).SetValue("部門C");
            sheet.SetExcelCellMerge(ExcelCol.C, ExcelCol.C, 8, 10);

            // Region merge example
            sheet.SetCellPosition(ExcelCol.A, 12).SetValue("重要通知");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.E, 12, 14);
            sheet.SetCellPosition(ExcelCol.A, 12).SetCellStyle("HighlightYellow");

            Console.WriteLine("  ✓ CellMergeExample 建立完成");
        }

        /// <summary>
        /// Example 8: Insert picture example
        /// </summary>
        static void CreatePictureExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 PictureExample...");

            var sheet = fluent.UseSheet("PictureExample", true);
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.D, 20);

            var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "pain.jpg");

            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Warning: Image file not found: {imagePath}");
                return;
            }

            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // 1. Basic picture insertion
            sheet.SetCellPosition(ExcelCol.A, 1)
                .SetValue("基本插入（自動高度）")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.A, 2)
                .SetPictureOnCell(imageBytes, 200);

            // 2. Manually set width and height
            sheet.SetCellPosition(ExcelCol.B, 1)
                .SetValue("手動設置尺寸")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.B, 2)
                .SetPictureOnCell(imageBytes, 200, 150);

            // 3. Custom column width conversion ratio
            sheet.SetCellPosition(ExcelCol.C, 1)
                .SetValue("自定義列寬比例")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.C, 2)
                .SetPictureOnCell(imageBytes, 300, AnchorType.MoveAndResize, 5.0);

            // 4. Chained call example
            sheet.SetCellPosition(ExcelCol.D, 1)
                .SetValue("鏈式調用示例")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.D, 2)
                .SetPictureOnCell(imageBytes, 180, 180, AnchorType.MoveAndResize, 7.0)
                .SetCellPosition(ExcelCol.D, 10)
                .SetValue("圖片下方文字");

            // 5. Different anchor type examples
            sheet.SetCellPosition(ExcelCol.A, 5)
                .SetValue("MoveAndResize（默認）")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.A, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.MoveAndResize);

            sheet.SetCellPosition(ExcelCol.B, 5)
                .SetValue("MoveDontResize")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.B, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.MoveDontResize);

            sheet.SetCellPosition(ExcelCol.C, 5)
                .SetValue("DontMoveAndResize")
                .SetCellStyle("HeaderBlue");

            sheet.SetCellPosition(ExcelCol.C, 6)
                .SetPictureOnCell(imageBytes, 150, 150, AnchorType.DontMoveAndResize);

            // 6. Multiple pictures arrangement
            sheet.SetCellPosition(ExcelCol.A, 9)
                .SetValue("多圖片排列")
                .SetCellStyle("HeaderBlue");

            for (int i = 0; i < 3; i++)
            {
                sheet.SetCellPosition((ExcelCol)((int)ExcelCol.A + i), 10)
                    .SetPictureOnCell(imageBytes, 100, 100);
            }

            Console.WriteLine("  ✓ PictureExample 建立完成");
        }

        #endregion
    }
}
