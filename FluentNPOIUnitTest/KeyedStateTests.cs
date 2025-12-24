using System;
using System.IO;
using System.Linq;
using Xunit;
using FluentNPOI.HotReload.State;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Phase 4: Keyed State Management
    /// </summary>
    public class KeyedStateTests
    {
        #region StateTypes Tests

        [Fact]
        public void CellLocation_ShouldFormatCorrectly()
        {
            // Arrange
            var location = new CellLocation("Sheet1", 5, 2);

            // Act
            var result = location.ToString();

            // Assert
            Assert.Equal("Sheet1!C5", result);
        }

        [Fact]
        public void StateEntry_ShouldStoreValues()
        {
            // Arrange & Act
            var entry = new StateEntry
            {
                Key = "test_key",
                Value = 42,
                Location = new CellLocation("Sheet1", 1, 0)
            };

            // Assert
            Assert.Equal("test_key", entry.Key);
            Assert.Equal(42, entry.Value);
            Assert.NotNull(entry.Location);
            Assert.True(entry.LastUpdated <= DateTime.Now);
        }

        [Fact]
        public void StateStorageStrategy_ShouldHaveExpectedValues()
        {
            // Assert
            Assert.Equal(0, (int)StateStorageStrategy.Comments);
            Assert.Equal(1, (int)StateStorageStrategy.CustomXmlParts);
            Assert.Equal(2, (int)StateStorageStrategy.HiddenSheet);
            Assert.Equal(3, (int)StateStorageStrategy.None);
        }

        #endregion

        #region KeyedStateManager Basic Tests

        [Fact]
        public void KeyedStateManager_SetValue_ShouldStoreValue()
        {
            // Arrange
            var manager = new KeyedStateManager();

            // Act
            manager.SetValue("key1", "value1");
            manager.SetValue("key2", 42);

            // Assert
            Assert.Equal(2, manager.CachedEntryCount);
        }

        [Fact]
        public void KeyedStateManager_GetValue_ShouldRetrieveValue()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("name", "John");
            manager.SetValue("age", 30);

            // Act & Assert
            Assert.Equal("John", manager.GetValue<string>("name"));
            Assert.Equal(30, manager.GetValue<int>("age"));
        }

        [Fact]
        public void KeyedStateManager_GetValue_ShouldReturnDefaultWhenNotFound()
        {
            // Arrange
            var manager = new KeyedStateManager();

            // Act
            var result = manager.GetValue("missing", "default");

            // Assert
            Assert.Equal("default", result);
        }

        [Fact]
        public void KeyedStateManager_HasKey_ShouldReturnCorrectly()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("exists", 1);

            // Assert
            Assert.True(manager.HasKey("exists"));
            Assert.False(manager.HasKey("missing"));
            Assert.False(manager.HasKey(""));
            Assert.False(manager.HasKey(null!));
        }

        [Fact]
        public void KeyedStateManager_Clear_ShouldRemoveAllEntries()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("key1", "value1");
            manager.SetValue("key2", "value2");
            Assert.Equal(2, manager.CachedEntryCount);

            // Act
            manager.Clear();

            // Assert
            Assert.Equal(0, manager.CachedEntryCount);
        }

        [Fact]
        public void KeyedStateManager_GetAllEntries_ShouldReturnAllEntries()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("a", 1);
            manager.SetValue("b", 2);
            manager.SetValue("c", 3);

            // Act
            var entries = manager.GetAllEntries().ToList();

            // Assert
            Assert.Equal(3, entries.Count);
        }

        #endregion

        #region Storage Strategy Tests

        [Fact]
        public void KeyedStateManager_HiddenSheet_ShouldPersistState()
        {
            // Arrange
            var manager = new KeyedStateManager { StorageStrategy = StateStorageStrategy.HiddenSheet };
            manager.SetValue("product", "Apple");
            manager.SetValue("quantity", 100);

            var workbook = new XSSFWorkbook();

            // Act - Save
            manager.SaveToWorkbook(workbook);

            // Assert - Hidden sheet exists
            var stateSheet = workbook.GetSheet("__FluentNPOI_State__");
            Assert.NotNull(stateSheet);
            // Check sheet is hidden
            var sheetIndex = workbook.GetSheetIndex("__FluentNPOI_State__");
            Assert.True(workbook.IsSheetHidden(sheetIndex) || workbook.IsSheetVeryHidden(sheetIndex));
        }

        [Fact]
        public void KeyedStateManager_HiddenSheet_ShouldLoadState()
        {
            // Arrange - Create workbook with state
            var saveManager = new KeyedStateManager { StorageStrategy = StateStorageStrategy.HiddenSheet };
            saveManager.SetValue("test_key", "test_value");
            saveManager.SetValue("number", 42);

            var workbook = new XSSFWorkbook();
            saveManager.SaveToWorkbook(workbook);

            // Act - Load into new manager
            var loadManager = new KeyedStateManager { StorageStrategy = StateStorageStrategy.HiddenSheet };
            loadManager.LoadFromWorkbook(workbook);

            // Assert
            Assert.Equal("test_value", loadManager.GetValue<string>("test_key"));
            Assert.Equal(42.0, loadManager.GetValue<double>("number")); // NPOI stores as double
        }

        [Fact]
        public void KeyedStateManager_None_ShouldNotPersist()
        {
            // Arrange
            var manager = new KeyedStateManager { StorageStrategy = StateStorageStrategy.None };
            manager.SetValue("key", "value");

            var workbook = new XSSFWorkbook();

            // Act
            manager.SaveToWorkbook(workbook);

            // Assert - No hidden sheet created
            var stateSheet = workbook.GetSheet("__FluentNPOI_State__");
            Assert.Null(stateSheet);
        }

        #endregion

        #region GenerateKeyComment Tests

        [Fact]
        public void GenerateKeyComment_ShouldCreateCorrectFormat()
        {
            // Act
            var comment = KeyedStateManager.GenerateKeyComment("my_key");

            // Assert
            Assert.Equal("Key: my_key", comment);
        }

        [Fact]
        public void GenerateKeyComment_WithSourceInfo_ShouldIncludeIt()
        {
            // Act
            var comment = KeyedStateManager.GenerateKeyComment("my_key", "Source: Test.cs:42");

            // Assert
            Assert.Contains("Key: my_key", comment);
            Assert.Contains("Source: Test.cs:42", comment);
        }

        #endregion

        #region Type Conversion Tests

        [Fact]
        public void KeyedStateManager_GetValue_ShouldConvertTypes()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("double_as_int", 42.0);

            // Act - Try to get as int
            var result = manager.GetValue<int>("double_as_int", 0);

            // Assert
            Assert.Equal(42, result);
        }

        [Fact]
        public void KeyedStateManager_SetValue_ShouldOverwriteExisting()
        {
            // Arrange
            var manager = new KeyedStateManager();
            manager.SetValue("key", "old_value");

            // Act
            manager.SetValue("key", "new_value");

            // Assert
            Assert.Equal("new_value", manager.GetValue<string>("key"));
            Assert.Equal(1, manager.CachedEntryCount);
        }

        #endregion
    }
}
