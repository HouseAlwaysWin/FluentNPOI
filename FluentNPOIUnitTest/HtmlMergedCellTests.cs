using System;
using NPOI.XSSF.UserModel;
using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;

namespace FluentNPOIUnitTest
{
    public class HtmlMergedCellTests
    {
        [Fact]
        public void ToHtmlString_VerticalMerge_RendersRowspan()
        {
            var wb = new FluentWorkbook(new XSSFWorkbook());
            var sheet = wb.UseSheet("V", true);
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Tall");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 1, 3); // A1:A3

            var html = wb.ToHtmlString(false);

            Assert.Contains("rowspan=\"3\"", html);
            Assert.Contains("Tall", html);
        }

        [Fact]
        public void ToHtmlString_BlockMerge_RendersRowspanAndColspan()
        {
            var wb = new FluentWorkbook(new XSSFWorkbook());
            var sheet = wb.UseSheet("B", true);
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Block");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1, 3); // A1:C3

            var html = wb.ToHtmlString(false);

            Assert.Contains("rowspan=\"3\"", html);
            Assert.Contains("colspan=\"3\"", html);
            Assert.Contains("Block", html);
        }

        [Fact]
        public void ToHtmlString_MultipleMergedRegions_RendersEach()
        {
            var wb = new FluentWorkbook(new XSSFWorkbook());
            var sheet = wb.UseSheet("M", true);
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Row1");
            sheet.SetCellPosition(ExcelCol.A, 2).SetValue("Row2");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 1, 1); // A1:B1
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 2, 2); // A2:B2

            var html = wb.ToHtmlString(false);

            Assert.Equal(2, CountOccurrences(html, "colspan=\"2\""));
        }

        private static int CountOccurrences(string haystack, string needle)
        {
            int count = 0, idx = 0;
            while ((idx = haystack.IndexOf(needle, idx, StringComparison.Ordinal)) != -1)
            {
                count++;
                idx += needle.Length;
            }
            return count;
        }
    }
}
