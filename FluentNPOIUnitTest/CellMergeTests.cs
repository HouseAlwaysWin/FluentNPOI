using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using System;
using FluentNPOI.Stages;

namespace FluentNPOIUnitTest
{
    public class CellMergeTests
    {
        [Fact]
        public void SetExcelCellMerge_HorizontalMerge_ShouldMergeCellsInRow()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 横向合并 A1 到 C1
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // 1-based to 0-based: row 1 -> 0
            Assert.Equal(0, mergedRegion.LastRow);
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(2, mergedRegion.LastColumn); // C -> 2
        }

        [Fact]
        public void SetExcelCellMerge_VerticalMerge_ShouldMergeCellsInColumn()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 纵向合并 A1 到 A5
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 1, 5);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(4, mergedRegion.LastRow); // row 5 -> 4
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(0, mergedRegion.LastColumn); // A -> 0
        }

        [Fact]
        public void SetExcelCellMerge_RegionMerge_ShouldMergeCellRegion()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 区域合并 A1 到 C3
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1, 3);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(2, mergedRegion.LastRow); // row 3 -> 2
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(2, mergedRegion.LastColumn); // C -> 2
        }

        [Fact]
        public void SetExcelCellMerge_MultipleMerges_ShouldCreateMultipleRegions()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 创建多个合并区域
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1) // 横向合并第1行
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 2, 5) // 纵向合并 A2-A5
                .SetExcelCellMerge(ExcelCol.D, ExcelCol.F, 1, 3); // 区域合并 D1-F3

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(3, sheet.NumMergedRegions);

            // 验证第一个合并区域（横向）
            var region1 = sheet.GetMergedRegion(0);
            Assert.Equal(0, region1.FirstRow);
            Assert.Equal(0, region1.LastRow);
            Assert.Equal(0, region1.FirstColumn); // A
            Assert.Equal(2, region1.LastColumn); // C

            // 验证第二个合并区域（纵向）
            var region2 = sheet.GetMergedRegion(1);
            Assert.Equal(1, region2.FirstRow); // row 2 -> 1
            Assert.Equal(4, region2.LastRow); // row 5 -> 4
            Assert.Equal(0, region2.FirstColumn); // A
            Assert.Equal(0, region2.LastColumn); // A

            // 验证第三个合并区域（区域）
            var region3 = sheet.GetMergedRegion(2);
            Assert.Equal(0, region3.FirstRow); // row 1 -> 0
            Assert.Equal(2, region3.LastRow); // row 3 -> 2
            Assert.Equal(3, region3.FirstColumn); // D -> 3
            Assert.Equal(5, region3.LastColumn); // F -> 5
        }

        [Fact]
        public void SetExcelCellMerge_ChainedCalls_ShouldReturnFluentSheet()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 链式调用
            var result = fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 1)
                .SetExcelCellMerge(ExcelCol.C, ExcelCol.D, 2);

            // Assert
            Assert.NotNull(result);
            Assert.IsType<FluentSheet>(result);
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(2, sheet.NumMergedRegions);
        }

        [Fact]
        public void SetExcelCellMerge_SingleCell_ShouldThrowException()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act & Assert - NPOI 不允许合并单个单元格，必须至少包含2个单元格
            Assert.Throws<ArgumentException>(() =>
            {
                fluentWorkbook.UseSheet("Sheet1")
                    .SetExcelCellMerge(ExcelCol.A, ExcelCol.A, 1);
            });
        }

        [Fact]
        public void SetExcelCellMerge_WithCellValues_ShouldPreserveFirstCellValue()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 先设置值，再合并
            sheet.SetCellPosition(ExcelCol.A, 1).SetValue("Merged Title");
            sheet.SetCellPosition(ExcelCol.B, 1).SetValue("Will be merged");
            sheet.SetCellPosition(ExcelCol.C, 1).SetValue("Also merged");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1);

            // Assert
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, npoiSheet.NumMergedRegions);

            // 验证合并后，只有第一个单元格有值
            var firstCell = npoiSheet.GetRow(0)?.GetCell(0);
            Assert.NotNull(firstCell);
            Assert.Equal("Merged Title", firstCell.StringCellValue);

            // 合并区域内的其他单元格应该为空或引用第一个单元格
            var mergedRegion = npoiSheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow);
            Assert.Equal(0, mergedRegion.LastRow);
            Assert.Equal(0, mergedRegion.FirstColumn);
            Assert.Equal(2, mergedRegion.LastColumn);
        }

        [Fact]
        public void SetExcelCellMerge_DifferentRows_ShouldMergeCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 合并不同行的区域
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.B, 1, 3); // A1-B3

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(2, mergedRegion.LastRow); // row 3 -> 2
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(1, mergedRegion.LastColumn); // B -> 1
        }

        [Fact]
        public void SetExcelCellMerge_LargeRegion_ShouldHandleCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);

            // Act - 合并大区域 A1 到 Z10
            fluentWorkbook.UseSheet("Sheet1")
                .SetExcelCellMerge(ExcelCol.A, ExcelCol.Z, 1, 10);

            // Assert
            var sheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, sheet.NumMergedRegions);

            var mergedRegion = sheet.GetMergedRegion(0);
            Assert.Equal(0, mergedRegion.FirstRow); // row 1 -> 0
            Assert.Equal(9, mergedRegion.LastRow); // row 10 -> 9
            Assert.Equal(0, mergedRegion.FirstColumn); // A -> 0
            Assert.Equal(25, mergedRegion.LastColumn); // Z -> 25
        }

        [Fact]
        public void SetExcelCellMerge_OverlappingRegions_ShouldThrowException()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 创建第一个合并区域
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.C, 1); // A1-C1

            // Assert - NPOI 不允许重叠的合并区域，应该抛出异常
            Assert.Throws<InvalidOperationException>(() =>
            {
                sheet.SetExcelCellMerge(ExcelCol.B, ExcelCol.D, 1); // B1-D1 (与第一个重叠)
            });

            // 验证第一个合并区域已创建
            var npoiSheet = workbook.GetSheet("Sheet1");
            Assert.Equal(1, npoiSheet.NumMergedRegions);
            var region1 = npoiSheet.GetMergedRegion(0);
            Assert.Equal(0, region1.FirstColumn); // A
            Assert.Equal(2, region1.LastColumn); // C
        }
    }
}

