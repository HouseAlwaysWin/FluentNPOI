using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using FluentNPOI.Stages;

namespace NPOIPlusUnitTest
{
    public class GetTableColumnRangeTests
    {
        private class TestData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public DateTime DateOfBirth { get; set; }
            public bool IsActive { get; set; }
            public double Score { get; set; }
            public decimal Amount { get; set; }
        }

        private class PartialData
        {
            public int ID { get; set; }
            public string Name { get; set; } = string.Empty;
            public DateTime DateOfBirth { get; set; }
        }

        private class MiddleColumnsData
        {
            public bool IsActive { get; set; }
            public double Score { get; set; }
            public decimal Amount { get; set; }
        }

        private class SingleColumnData
        {
            public string Name { get; set; } = string.Empty;
        }

        [Fact]
        public void GetTable_WithColumnRange_ShouldReadOnlySpecifiedColumns()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m },
                new TestData { ID = 3, Name = "Charlie", DateOfBirth = new DateTime(1992, 3, 3), IsActive = true, Score = 92.0, Amount = 3000m }
            };

            // 先寫入完整數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 A-C 列（ID, Name, DateOfBirth）
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 4);

            // Assert
            Assert.Equal(3, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal(new DateTime(1990, 1, 1), readData[0].DateOfBirth);
            Assert.Equal(2, readData[1].ID);
            Assert.Equal("Bob", readData[1].Name);
            Assert.Equal(3, readData[2].ID);
            Assert.Equal("Charlie", readData[2].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_MiddleColumns_ShouldReadCorrectColumns()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            // 先寫入完整數據
            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 D-F 列（IsActive, Score, Amount）
            var readData = sheet.GetTable<MiddleColumnsData>(ExcelCol.D, ExcelCol.F, 2, 3);

            // Assert
            Assert.Equal(2, readData.Count);
            Assert.True(readData[0].IsActive);
            Assert.Equal(95.5, readData[0].Score);
            Assert.Equal(1000m, readData[0].Amount);
            Assert.False(readData[1].IsActive);
            Assert.Equal(87.0, readData[1].Score);
            Assert.Equal(2000m, readData[1].Amount);
        }

        [Fact]
        public void GetTable_WithColumnRange_SingleColumn_ShouldReadOnlyOneColumn()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 B 列（Name）
            var readData = sheet.GetTable<SingleColumnData>(ExcelCol.B, ExcelCol.B, 2, 3);

            // Assert
            Assert.Equal(2, readData.Count);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal("Bob", readData[1].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_MoreColumnsThanProperties_ShouldUseOnlyAvailableProperties()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 指定 A-F 列，但 PartialData 只有3個屬性
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.F, 2, 2);

            // Assert - 應該只使用前3個屬性（ID, Name, DateOfBirth）
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            Assert.Equal(new DateTime(1990, 1, 1), readData[0].DateOfBirth);
        }

        [Fact]
        public void GetTable_WithColumnRange_FewerColumnsThanProperties_ShouldLeaveRemainingPropertiesDefault()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取 A-B 列，但 TestData 有6個屬性
            var readData = sheet.GetTable<TestData>(ExcelCol.A, ExcelCol.B, 2, 2);

            // Assert - 只有前2個屬性會被填充
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
            // 其他屬性應該保持默認值
            Assert.Equal(default(DateTime), readData[0].DateOfBirth);
            Assert.False(readData[0].IsActive);
            Assert.Equal(0.0, readData[0].Score);
            Assert.Equal(0m, readData[0].Amount);
        }

        [Fact]
        public void GetTable_WithColumnRange_EmptyRows_ShouldSkipEmptyRows()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m },
                new TestData { ID = 2, Name = "Bob", DateOfBirth = new DateTime(1991, 2, 2), IsActive = false, Score = 87.0, Amount = 2000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            // 創建空行
            var npoiSheet = workbook.GetSheet("Sheet1");
            npoiSheet.CreateRow(3); // 第4行（0-based index 3）

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 讀取第2-5行，但第4行是空的
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 5);

            // Assert - 應該只讀取2行數據，跳過空行
            Assert.Equal(2, readData.Count);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal(2, readData[1].ID);
        }

        [Fact]
        public void GetTable_WithColumnRange_SingleRow_ShouldReadCorrectly()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - 只讀取單一行
            var readData = sheet.GetTable<PartialData>(ExcelCol.A, ExcelCol.C, 2, 2);

            // Assert
            Assert.Single(readData);
            Assert.Equal(1, readData[0].ID);
            Assert.Equal("Alice", readData[0].Name);
        }

        [Fact]
        public void GetTable_WithColumnRange_StartColAfterEndCol_ShouldHandleGracefully()
        {
            // Arrange
            var workbook = new XSSFWorkbook();
            var fluentWorkbook = new FluentWorkbook(workbook);
            var testData = new List<TestData>
            {
                new TestData { ID = 1, Name = "Alice", DateOfBirth = new DateTime(1990, 1, 1), IsActive = true, Score = 95.5, Amount = 1000m }
            };

            fluentWorkbook.UseSheet("Sheet1")
                .SetTable(testData, ExcelCol.A, 1)
                .BeginTitleSet("ID")
                .BeginBodySet("ID").End()
                .BeginTitleSet("Name")
                .BeginBodySet("Name").End()
                .BeginTitleSet("DateOfBirth")
                .BeginBodySet("DateOfBirth").End()
                .BeginTitleSet("IsActive")
                .BeginBodySet("IsActive").End()
                .BeginTitleSet("Score")
                .BeginBodySet("Score").End()
                .BeginTitleSet("Amount")
                .BeginBodySet("Amount").End()
                .BuildRows();

            var sheet = fluentWorkbook.UseSheet("Sheet1");

            // Act - startCol > endCol 的情況（雖然不合理，但應該能處理）
            var readData = sheet.GetTable<PartialData>(ExcelCol.C, ExcelCol.A, 2, 2);

            // Assert - 應該返回空列表或處理錯誤（實際行為取決於實現）
            // 由於 columnCount 會是負數，membersToUse 會是 0，所以不會讀取任何數據
            Assert.Empty(readData);
        }
    }
}

