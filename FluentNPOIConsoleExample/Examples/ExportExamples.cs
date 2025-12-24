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
            Console.WriteLine("Creating HtmlExportExample...");

            var htmlPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedRequest.html";

            Console.WriteLine("  > Creating 'HtmlDemo' Sheet to demonstrate style support...");

            // Define Styles
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

            // Create Sheet Content
            var sheet = fluent.UseSheet("HtmlDemo", true);

            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("HTML Export Feature Demo")
                 .SetCellStyle("MergedTitle");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.D, 1, 1);

            // Row 2: Color Demo
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("Red Background").SetCellStyle("RedBg");
            sheet.SetCellPosition(ExcelCol.B, 2).SetValue("Green Italic Text").SetCellStyle("GreenText");
            sheet.SetCellPosition(ExcelCol.C, 2).SetValue(1234.56).SetCellStyle("NumberFmt");
            sheet.SetCellPosition(ExcelCol.D, 2).SetValue(9999.99).SetCellStyle("Currency");

            // Row 3: Text Decoration Demo
            sheet.SetCellPosition(ExcelCol.A, 3).SetValue("Underlined Text").SetCellStyle("Underline");
            sheet.SetCellPosition(ExcelCol.B, 3).SetValue("Strikethrough Text").SetCellStyle("Strikethrough");
            sheet.SetCellPosition(ExcelCol.C, 3).SetValue("Plain Text");
            sheet.SetCellPosition(ExcelCol.D, 3).SetValue(0.1234).SetCellStyle("NumberFmt");

            // Row 4-6: Merged Cells Demo
            sheet.SetCellPosition(ExcelCol.A, 4).SetValue("Vertical\nMerge").SetCellStyle("MergedTitle");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 4, 6);

            sheet.SetCellPosition(ExcelCol.B, 4).SetValue("2x2 Block Merge").SetCellStyle("RedBg");
            sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.C, 4, 5);

            sheet.SetCellPosition(ExcelCol.D, 4).SetValue("D4");
            sheet.SetCellPosition(ExcelCol.D, 5).SetValue("D5");
            sheet.SetCellPosition(ExcelCol.B, 6).SetValue("B6");
            sheet.SetCellPosition(ExcelCol.C, 6).SetValue("C6");
            sheet.SetCellPosition(ExcelCol.D, 6).SetValue("D6");

            // Export to HTML
            fluent.SaveAsHtml(htmlPath, fullHtml: true);
            var htmlFragment = fluent.ToHtmlString(fullHtml: false);

            Console.WriteLine($"  ✓ HTML Export Completed: {htmlPath}");
            Console.WriteLine($"  ✓ HTML Fragment Preview (First 100 chars): {htmlFragment.Substring(0, Math.Min(100, htmlFragment.Length))}...");
        }

        /// <summary>
        /// Example 13: Export to PDF with merged cells
        /// </summary>
        public static void CreatePdfExportExample(FluentWorkbook fluent)
        {
            Console.WriteLine("Creating PdfExportExample...");

            var pdfPath = @$"{AppDomain.CurrentDomain.BaseDirectory}\Resources\ExportedReport.pdf";

            var sheet = fluent.UseSheet("HtmlDemo", true);

            FluentNPOI.Pdf.PdfConverter.ConvertSheetToPdf(sheet.GetSheet(), fluent.GetWorkbook(), pdfPath);

            Console.WriteLine($"  ✓ PDF Export Completed: {pdfPath}");
            Console.WriteLine("  > PDF Support: Background Color, Text Color, Bold/Italic, Underline/Strikeout,");
            Console.WriteLine("              Border Styles, Number Formatting, Text Alignment, Merged Cells");
        }

        #endregion
    }
}
