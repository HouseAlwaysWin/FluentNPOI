using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace FluentNPOI.Html
{
    /// <summary>
    /// Helper class to convert Excel Sheet to HTML
    /// </summary>
    public static class HtmlConverter
    {
        public static string ConvertSheetToHtml(ISheet sheet, bool fullHtml = true)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            var sb = new StringBuilder();
            var styles = new Dictionary<int, string>(); // Index -> CSS Class Name
            var cssRules = new StringBuilder();

            // 1. Analyze Merged Regions
            var mergeMap = BuildMergeMap(sheet);

            // 2. Build HTML Body
            var body = BuildHtmlBody(sheet, mergeMap, styles);

            // 3. Build CSS
            foreach (var kvp in styles)
            {
                var styleIndex = kvp.Key;
                var className = kvp.Value;
                var style = sheet.Workbook.GetCellStyleAt((short)styleIndex);
                var css = GenerateCss(style, sheet.Workbook);
                cssRules.AppendLine($".{className} {{ {css} }}");
            }

            // 4. Assemble final HTML
            if (fullHtml)
            {
                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style>");
                sb.AppendLine("table { border-collapse: collapse; width: 100%; }");
                sb.AppendLine("td, th { padding: 4px; border: 1px solid #ddd; }");
                sb.Append(cssRules);
                sb.AppendLine("</style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.Append(body);
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");
            }
            else
            {
                sb.AppendLine("<style>");
                sb.Append(cssRules);
                sb.AppendLine("</style>");
                sb.Append(body);
            }

            return sb.ToString();
        }

        private static string BuildHtmlBody(ISheet sheet, Dictionary<string, CellRangeAddress> mergeMap, Dictionary<int, string> styles)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<table>");

            var rowEnumerator = sheet.GetRowEnumerator();
            while (rowEnumerator.MoveNext())
            {
                var row = (IRow)rowEnumerator.Current;
                if (row == null) continue;

                sb.AppendLine("<tr>");

                var lastCellNum = row.LastCellNum;
                for (int i = 0; i < lastCellNum; i++)
                {
                    var cell = row.GetCell(i);
                    var key = $"{row.RowNum}_{i}";

                    // Check if this cell is inside a merged region
                    if (IsInsideMergedRegion(row.RowNum, i, mergeMap, out var region))
                    {
                        // If it's the top-left cell of the region, render it with rowspan/colspan
                        if (region.FirstRow == row.RowNum && region.FirstColumn == i)
                        {
                            var rowSpan = region.LastRow - region.FirstRow + 1;
                            var colSpan = region.LastColumn - region.FirstColumn + 1;
                            RenderCell(sb, cell, styles, rowSpan, colSpan);
                        }
                        // Else: Skip (it's covered by the spanning cell)
                    }
                    else
                    {
                        // Normal cell
                        RenderCell(sb, cell, styles);
                    }
                }

                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");
            return sb.ToString();
        }

        private static void RenderCell(StringBuilder sb, ICell cell, Dictionary<int, string> styles, int rowSpan = 1, int colSpan = 1)
        {
            sb.Append("<td");

            if (rowSpan > 1) sb.Append($" rowspan=\"{rowSpan}\"");
            if (colSpan > 1) sb.Append($" colspan=\"{colSpan}\"");

            if (cell != null)
            {
                // Register Style
                var styleIdx = (int)cell.CellStyle.Index;
                if (!styles.ContainsKey(styleIdx))
                {
                    styles[styleIdx] = $"s{styleIdx}";
                }
                sb.Append($" class=\"{styles[styleIdx]}\"");
            }

            sb.Append(">");

            // Content
            if (cell != null)
            {
                sb.Append(GetCellValueHtml(cell));
            }

            sb.AppendLine("</td>");
        }

        private static string GetCellValueHtml(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return System.Net.WebUtility.HtmlEncode(cell.RichStringCellValue?.String ?? "");
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return string.Format("{0:yyyy-MM-dd}", cell.DateCellValue);
                    }
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.Formula:
                    // Evaluate formula? For now just show cached value
                    try
                    {
                        if (cell.CachedFormulaResultType == CellType.Numeric) return cell.NumericCellValue.ToString();
                        if (cell.CachedFormulaResultType == CellType.String) return cell.StringCellValue;
                    }
                    catch { }
                    return "=" + cell.CellFormula;
                default:
                    return "";
            }
        }

        private static bool IsInsideMergedRegion(int row, int col, Dictionary<string, CellRangeAddress> map, out CellRangeAddress region)
        {
            // Naive check: iterate map? No, map should be keyed by cell coordinates?
            // Actually, for O(1) check we need a 2D map or dictionary of every coordinate.
            // Or just iterate implementation since merged regions list is usually small.
            // Let's rely on the pre-built map of "TopLeft" -> Region is not enough.
            // We need "AnyCoordinate" -> Region.

            // Optimization: Let's assume BuildMergeMap returns a dictionary where Key = "row_col" for ALL cells in region.
            var key = $"{row}_{col}";
            return map.TryGetValue(key, out region);
        }

        private static Dictionary<string, CellRangeAddress> BuildMergeMap(ISheet sheet)
        {
            var map = new Dictionary<string, CellRangeAddress>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                for (int r = region.FirstRow; r <= region.LastRow; r++)
                {
                    for (int c = region.FirstColumn; c <= region.LastColumn; c++)
                    {
                        map[$"{r}_{c}"] = region;
                    }
                }
            }
            return map;
        }

        private static string GenerateCss(ICellStyle style, IWorkbook wb)
        {
            var sb = new StringBuilder();

            // Background Color (FillForegroundColor)
            // Note: In NPOI, FillForegroundColor store the color index.
            if (style.FillPattern == FillPattern.SolidForeground)
            {
                var hex = GetColorHex(style.FillForegroundColorColor, wb);
                if (!string.IsNullOrEmpty(hex)) sb.Append($"background-color: {hex}; ");
            }

            // Borders
            AppendBorder(sb, "top", style.BorderTop, style.TopBorderColor, wb);
            AppendBorder(sb, "right", style.BorderRight, style.RightBorderColor, wb);
            AppendBorder(sb, "bottom", style.BorderBottom, style.BottomBorderColor, wb);
            AppendBorder(sb, "left", style.BorderLeft, style.LeftBorderColor, wb);

            // Alignment
            if (style.Alignment == HorizontalAlignment.Center) sb.Append("text-align: center; ");
            if (style.Alignment == HorizontalAlignment.Right) sb.Append("text-align: right; ");

            // Font
            var font = style.GetFont(wb);
            if (font != null)
            {
                // Font color
                // For HSSF, font.Color is short index.
                // For XSSF, we need to cast to XSSFFont to get XSSFColor if possible, or use IFont color index
                var hex = GetFontColorHex(font, wb);
                if (!string.IsNullOrEmpty(hex)) sb.Append($"color: {hex}; ");

                if (font.IsBold) sb.Append("font-weight: bold; ");
                if (font.IsItalic) sb.Append("font-style: italic; ");
                if (font.IsStrikeout) sb.Append("text-decoration: line-through; ");
                if (font.Underline != FontUnderlineType.None) sb.Append("text-decoration: underline; ");

                if (font.FontHeightInPoints > 0) sb.Append($"font-size: {font.FontHeightInPoints}pt; ");
                if (!string.IsNullOrEmpty(font.FontName)) sb.Append($"font-family: '{font.FontName}'; ");
            }

            return sb.ToString();
        }

        private static void AppendBorder(StringBuilder sb, string side, BorderStyle border, short colorIndex, IWorkbook wb)
        {
            if (border == BorderStyle.None) return;

            string style = "solid";
            string width = "1px";

            switch (border)
            {
                case BorderStyle.Thin:
                case BorderStyle.Hair:
                    width = "1px";
                    break;
                case BorderStyle.Medium:
                    width = "2px";
                    break;
                case BorderStyle.Thick:
                    width = "3px";
                    break;
                case BorderStyle.Dashed:
                case BorderStyle.DashDot:
                case BorderStyle.DashDotDot:
                case BorderStyle.MediumDashed:
                case BorderStyle.MediumDashDot:
                case BorderStyle.MediumDashDotDot:
                case BorderStyle.SlantedDashDot:
                    style = "dashed";
                    break;
                case BorderStyle.Dotted:
                    style = "dotted";
                    break;
                case BorderStyle.Double:
                    style = "double";
                    width = "3px";
                    break;
            }

            var colorHex = GetColorFromIndex(colorIndex, wb) ?? "#000000"; // Default to black
            sb.Append($"border-{side}: {width} {style} {colorHex}; ");
        }

        private static string GetFontColorHex(IFont font, IWorkbook wb)
        {
            if (font is NPOI.XSSF.UserModel.XSSFFont xFont)
            {
                var xColor = xFont.GetXSSFColor();
                if (xColor != null) return GetXSSFColorHex(xColor);
            }

            // Fallback to index (works for HSSF and XSSF if indexed)
            return GetColorFromIndex(font.Color, wb);
        }

        private static string GetColorHex(IColor color, IWorkbook wb)
        {
            if (color == null) return null;

            // XSSF
            if (color is NPOI.XSSF.UserModel.XSSFColor xColor)
            {
                return GetXSSFColorHex(xColor);
            }

            // HSSF does not implement IColor well in all versions, often uses index.
            // But FillForegroundColorColor returns IColor which might be HSSFColor.
            if (color is NPOI.HSSF.Util.HSSFColor hColor)
            {
                var rgb = hColor.RGB;
                return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
            }

            return null;
        }

        private static string GetColorFromIndex(short index, IWorkbook wb)
        {
            if (index == IndexedColors.Automatic.Index || index == 0) return null;

            if (wb is NPOI.HSSF.UserModel.HSSFWorkbook hVal)
            {
                var palette = hVal.GetCustomPalette();
                var color = palette.GetColor(index);
                if (color != null)
                {
                    var rgb = color.RGB;
                    return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
                }
            }

            // XSSF Fallback: Use built-in map for standard colors
            if (StandardColorMap.TryGetValue(index, out var hex))
            {
                return hex;
            }

            return null;
        }

        private static string GetXSSFColorHex(NPOI.XSSF.UserModel.XSSFColor xColor)
        {
            if (xColor.IsRGB)
            {
                var rgb = xColor.RGB;
                if (rgb != null && rgb.Length == 3) // Sometimes it's ARGB with alpha? NPOI 2.7.1 usually RGB
                    return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
                if (rgb != null && rgb.Length == 4) // Alpha present
                    return $"#{rgb[1]:X2}{rgb[2]:X2}{rgb[3]:X2}"; // Skip alpha for CSS or use rgba?
            }
            // If XSSFColor works by index but IsRGB is false
            if (xColor.Index != 0)
            {
                if (StandardColorMap.TryGetValue(xColor.Index, out var hex)) return hex;
            }
            return null;
        }

        // Mapping from NPOI IndexedColors to Hex
        private static readonly Dictionary<short, string> StandardColorMap = new Dictionary<short, string>
        {
            { IndexedColors.Black.Index, "#000000" },
            { IndexedColors.White.Index, "#FFFFFF" },
            { IndexedColors.Red.Index, "#FF0000" },
            { IndexedColors.BrightGreen.Index, "#00FF00" },
            { IndexedColors.Blue.Index, "#0000FF" },
            { IndexedColors.Yellow.Index, "#FFFF00" },
            { IndexedColors.Pink.Index, "#FF00FF" },
            { IndexedColors.Turquoise.Index, "#00FFFF" },
            { IndexedColors.DarkRed.Index, "#800000" },
            { IndexedColors.Green.Index, "#008000" },
            { IndexedColors.DarkBlue.Index, "#000080" },
            { IndexedColors.DarkYellow.Index, "#808000" },
            { IndexedColors.Violet.Index, "#800080" },
            { IndexedColors.Teal.Index, "#008080" },
            { IndexedColors.Grey25Percent.Index, "#C0C0C0" },
            { IndexedColors.Grey50Percent.Index, "#808080" },
            { IndexedColors.Grey80Percent.Index, "#333333" },
            { IndexedColors.CornflowerBlue.Index, "#99CCFF" }, // Approx
            { IndexedColors.LightCornflowerBlue.Index, "#CCCCFF" }, // Approx
            { IndexedColors.Maroon.Index, "#800000" },
            { IndexedColors.LemonChiffon.Index, "#FFFACD" },
            { IndexedColors.LightTurquoise.Index, "#AFEEEE" },
            { IndexedColors.Orchid.Index, "#DA70D6" },
            { IndexedColors.Coral.Index, "#FF7F50" },
            { IndexedColors.RoyalBlue.Index, "#4169E1" },
            { IndexedColors.LightBlue.Index, "#ADD8E6" },
            { IndexedColors.SkyBlue.Index, "#87CEEB" },
            { IndexedColors.LightGreen.Index, "#90EE90" },
            { IndexedColors.LightYellow.Index, "#FFFFE0" },
            { IndexedColors.PaleBlue.Index, "#AFEEEE" },
            { IndexedColors.Rose.Index, "#FFC0CB" },
            { IndexedColors.Lavender.Index, "#E6E6FA" },
            { IndexedColors.Tan.Index, "#D2B48C" },
            { IndexedColors.Aqua.Index, "#00FFFF" },
            { IndexedColors.Lime.Index, "#00FF00" },
            { IndexedColors.Gold.Index, "#FFD700" },
            { IndexedColors.Orange.Index, "#FFA500" },
            { IndexedColors.Brown.Index, "#A52A2A" },
            { IndexedColors.Plum.Index, "#DDA0DD" },
            { IndexedColors.Indigo.Index, "#4B0082" },
            { IndexedColors.Grey40Percent.Index, "#969696" },
            { IndexedColors.DarkTeal.Index, "#003366" },
            { IndexedColors.SeaGreen.Index, "#2E8B57" },
            { IndexedColors.DarkGreen.Index, "#006400" },
            { IndexedColors.OliveGreen.Index, "#333300" }, // Approx
        };
    }
}
