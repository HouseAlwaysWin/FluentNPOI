using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class StyleCachingTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public bool IsActive { get; set; }
        }

        [Fact]
        public void DynamicStyle_ShouldCacheByKey()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, IsActive = true },
                new TestData { ID = 2, IsActive = false },
                new TestData { ID = 3, IsActive = true },
                new TestData { ID = 4, IsActive = false }
            };

            int styleSetterCallCount = 0;

            // Act
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("狀態")
                .BeginBodySet("IsActive")
                .SetCellStyle((styleParams) =>
                {
                    if (styleParams.GetRowItem<TestData>().IsActive)
                    {
                        return new CellStyleConfig("ActiveStyle", style =>
                        {
                            styleSetterCallCount++;
                            style.FillPattern = FillPattern.SolidForeground;
                            style.FillForegroundColor = IndexedColors.Green.Index;
                        });
                    }
                    return new CellStyleConfig("InactiveStyle", style =>
                    {
                        styleSetterCallCount++;
                        style.FillPattern = FillPattern.SolidForeground;
                        style.FillForegroundColor = IndexedColors.Yellow.Index;
                    });
                })
                .End()
                .BuildRows();

            // Assert - 4 行數據，但只有 2 種樣式，所以 StyleSetter 應該只被調用 2 次
            Assert.Equal(2, styleSetterCallCount);
        }

        [Fact]
        public void SetupCellStyle_ShouldApplyStyle()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, IsActive = true },
                new TestData { ID = 2, IsActive = true }
            };

            // Act
            fluentWorkbook
                .SetupCellStyle("FixedStyle", (wb, style) =>
                {
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = IndexedColors.Blue.Index;
                })
                .UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").SetCellStyle("FixedStyle").End()
                .BuildRows();

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            var cell1 = sheet.GetRow(1)?.GetCell(0);
            var cell2 = sheet.GetRow(2)?.GetCell(0);

            // 驗證樣式已套用
            Assert.NotNull(cell1?.CellStyle);
            Assert.NotNull(cell2?.CellStyle);
            Assert.Equal(FillPattern.SolidForeground, cell1.CellStyle.FillPattern);
            Assert.Equal(FillPattern.SolidForeground, cell2.CellStyle.FillPattern);
            Assert.Equal(IndexedColors.Blue.Index, cell1.CellStyle.FillForegroundColor);
            Assert.Equal(IndexedColors.Blue.Index, cell2.CellStyle.FillForegroundColor);
        }
    }
}

