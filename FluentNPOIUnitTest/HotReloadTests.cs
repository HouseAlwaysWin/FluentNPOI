using System;
using System.IO;
using Xunit;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload;
using FluentNPOI.HotReload.HotReload;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Phase 2: Hot Reload Infrastructure
    /// </summary>
    public class HotReloadTests
    {
        #region HotReloadHandler Tests

        [Fact]
        public void HotReloadHandler_IsActive_ShouldBeFalseWhenNoSubscribers()
        {
            // Assert
            Assert.False(HotReloadHandler.IsActive);
        }

        [Fact]
        public void HotReloadHandler_RefreshRequested_ShouldBeInvoked()
        {
            // Arrange
            bool wasInvoked = false;
            Type[]? receivedTypes = null;

            void Handler(Type[]? types)
            {
                wasInvoked = true;
                receivedTypes = types;
            }

            HotReloadHandler.RefreshRequested += Handler;

            try
            {
                // Act
                HotReloadHandler.TriggerRefresh();

                // Assert
                Assert.True(wasInvoked);
                Assert.Null(receivedTypes); // TriggerRefresh passes null
            }
            finally
            {
                HotReloadHandler.RefreshRequested -= Handler;
            }
        }

        [Fact]
        public void HotReloadHandler_IsActive_ShouldBeTrueWithSubscribers()
        {
            // Arrange
            void Handler(Type[]? _) { }
            HotReloadHandler.RefreshRequested += Handler;

            try
            {
                // Assert
                Assert.True(HotReloadHandler.IsActive);
            }
            finally
            {
                HotReloadHandler.RefreshRequested -= Handler;
            }
        }

        #endregion

        #region FluentHotReloadSession Tests

        [Fact]
        public void FluentHotReloadSession_Constructor_ShouldStoreOutputPath()
        {
            // Arrange & Act
            using var session = new FluentHotReloadSession("test/output.xlsx", wb => { });

            // Assert
            Assert.Equal("test/output.xlsx", session.OutputPath);
        }

        [Fact]
        public void FluentHotReloadSession_Refresh_ShouldIncrementRefreshCount()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });

                // Act
                session.Refresh();
                session.Refresh();

                // Assert
                Assert.Equal(2, session.RefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_Refresh_ShouldCreateExcelFile()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1).SetValue("Test");
                });

                // Act
                session.Refresh();

                // Assert
                Assert.True(File.Exists(tempPath));
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshCompleted_ShouldBeInvoked()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            int? completedRefreshCount = null;

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });
                session.RefreshCompleted += (count) => completedRefreshCount = count;

                // Act
                session.Refresh();

                // Assert
                Assert.Equal(1, completedRefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshError_ShouldBeInvokedOnFailure()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            Exception? capturedError = null;

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => throw new InvalidOperationException("Test error"));
                session.RefreshError += (ex) => capturedError = ex;

                // Act
                session.Refresh();

                // Assert
                Assert.NotNull(capturedError);
                Assert.IsType<InvalidOperationException>(capturedError);
                Assert.Equal("Test error", capturedError.Message);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_Start_ShouldPerformInitialRefresh()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });

                // Act
                session.Start();
                session.Stop();

                // Assert
                Assert.Equal(1, session.RefreshCount);
                Assert.True(File.Exists(tempPath));
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        #endregion

        #region Integration Tests

        [Fact]
        public void FluentHotReloadSession_WithHotReloadHandler_ShouldRespondToEvents()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb => { });
                session.Start();

                var initialCount = session.RefreshCount;

                // Act - simulate hot reload
                HotReloadHandler.TriggerRefresh();

                // Assert
                Assert.Equal(initialCount + 1, session.RefreshCount);

                session.Stop();
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_UpdateLogic_ShouldReflectInOutput()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            var currentValue = "Initial";

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(FluentNPOI.Models.ExcelCol.A, 1).SetValue(currentValue);
                });

                // Act - First refresh
                session.Refresh();

                // Change value in memory (simulating code change logic)
                currentValue = "Updated";
                session.Refresh();

                // Assert - Read the file and verify updated value
                using var fs = File.OpenRead(tempPath);
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheetAt(0);
                var cell = sheet.GetRow(0)?.GetCell(0);

                Assert.Equal("Updated", cell?.StringCellValue);
                Assert.Equal(2, session.RefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        #endregion
    }
}
