using System;
using System.IO;
using FluentNPOI;
using NPOI.SS.UserModel;
using FluentNPOI.Models;
using FluentNPOI.Stages;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Export Examples - HTML, PDF
    /// </summary>
    internal partial class Program
    {
        #region Export Examples

        /// <summary>
        /// Example 12: Export to HTML
        /// </summary>
        public static void CreateHtmlExportExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 HtmlExportExample...");

            var htmlPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedRequest.html";

            Console.WriteLine("  > 正在建立 'HtmlDemo' Sheet 以展示樣式支援...");

            // 定義樣式
            fluent.SetupCellStyle("MergedTitle", (w, s) =>
            {
                s.SetAlignment(HorizontalAlignment.Center);
                s.SetFontInfo(w, fontFamily: "Arial", fontHeight: 16, isBold: true);
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.LightCornflowerBlue);
                s.SetBorderAllStyle(BorderStyle.Thick);
            });

            fluent.SetupCellStyle("RedBg", (w, s) =>
            {
                s.FillPattern = FillPattern.SolidForeground;
                s.SetCellFillForegroundColor(IndexedColors.Red);
                s.SetFontInfo(w, color: IndexedColors.White);
                s.SetAlignment(HorizontalAlignment.Center);
            });

            fluent.SetupCellStyle("GreenText", (w, s) =>
            {
                s.SetFontInfo(w, color: IndexedColors.Green, isItalic: true, isBold: true);
                s.SetBorderAllStyle(BorderStyle.Dotted);
            });

            fluent.SetupCellStyle("NumberFmt", (w, s) =>
            {
                s.SetDataFormat(w, "#,##0.00");
                s.SetAlignment(HorizontalAlignment.Right);
            });

            fluent.SetupCellStyle("Currency", (w, s) =>
            {
                s.SetDataFormat(w, "$#,##0.00");
                s.SetAlignment(HorizontalAlignment.Right);
                s.SetFontInfo(w, isBold: true);
            });

            fluent.SetupCellStyle("Underline", (w, s) =>
            {
                var font = w.CreateFont();
                font.Underline = FontUnderlineType.Single;
                s.SetFont(font);
            });

            fluent.SetupCellStyle("Strikethrough", (w, s) =>
            {
                s.SetFontInfo(w, isStrikeout: true);
            });

            // 建立 Sheet 內容
            var sheet = fluent.UseSheet("HtmlDemo", true);

            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("HTML Export Feature Demo")
                 .SetCellStyle("MergedTitle");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1, 1);

            // Row 2: 顏色示範
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("Red Background").SetCellStyle("RedBg");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("Green Italic Text").SetCellStyle("GreenText");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue(1234.56).SetCellStyle("NumberFmt");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue(9999.99).SetCellStyle("Currency");

            // Row 3: 文字裝飾示範
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("Underlined Text").SetCellStyle("Underline");
            sheet.SetCellPosition(ExcelCol.B, 3).SetValue("Strikethrough Text").SetCellStyle("Strikethrough");
            sheet.SetCellPosition(ExcelCol.C, 3).SetValue("Plain Text");
            sheet.SetCellPosition(ExcelCol.D, 3).SetValue(0.1234).SetCellStyle("NumberFmt");

            // Row 4-6: 合併儲存格示範
            sheet.SetCellPosition(ExcelCol.A, 4).SetValue("Vertical\nMerge").SetCellStyle("MergedTitle");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 4, 6);

            sheet.SetCellPosition(ExcelCol.B, 4).SetValue("2x2 Block Merge").SetCellStyle("RedBg");
            sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.C, 4, 5);

            sheet.SetCellPosition(ExcelCol.D, 4).SetValue("D4");
            sheet.SetCellPosition(ExcelCol.D, 5).SetValue("D5");
            sheet.SetCellPosition(ExcelCol.B, 6).SetValue("B6");
            sheet.SetCellPosition(ExcelCol.C, 6).SetValue("C6");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue("D6");

            // 匯出為 HTML
            fluent.SaveAsHtml(htmlPath, fullHtml: true);
            var htmlFragment = fluent.ToHtmlString(fullHtml: false);

            Console.WriteLine($"  ✓ HTML 匯出完成: {htmlPath}");
            Console.WriteLine($"  ✓ HTML 片段預覽 (前 100 字): {htmlFragment.Substring(0, Math.Min(100, htmlFragment.Length))}...");
        }

        /// <summary>
        /// Example 13: Export to PDF with merged cells
        /// </summary>
        public static void CreatePdfExportExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 PdfExportExample...");

            var pdfPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedReport.pdf";

            var sheet = fluent.UseSheet("HtmlDemo", true);

            FluentNPOI.Pdf.PdfConverter.ConvertSheetToPdf(sheet.GetSheet(), fluent.GetWorkbook(), pdfPath);

            Console.WriteLine($"  ✓ PDF 匯出完成: {pdfPath}");
            Console.WriteLine("  > PDF 支援: 背景色、文字顏色、粗體/斜體、底線/刪除線、");
            Console.WriteLine("              邊框樣式、數值格式化、文字對齊、合併儲存格");
        }

        #endregion
    }
}
