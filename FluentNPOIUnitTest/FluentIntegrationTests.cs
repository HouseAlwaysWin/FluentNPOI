using System;
using System.IO;
using Xunit;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.HotReload;
using NPOI.XSSF.UserModel;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI and HotReload integration.
    /// Verifies that existing FluentNPOI code can work with the hot reload system.
    /// </summary>
    public class FluentIntegrationTests
    {
        #region FluentHotReloadSession Tests

        [Fact]
        public void FluentHotReloadSession_ShouldBuildWithFluentCode()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                Action<FluentWorkbook> buildAction = wb =>
                {
                    wb.UseSheet("Sheet1")
                        .SetCellPosition(ExcelCol.A, 1)
                        .SetValue("FluentNPOI Works!")
                        .SetCellPosition(ExcelCol.B, 1)
                        .SetValue(123);
                };

                using var session = new FluentHotReloadSession(tempPath, buildAction);

                // Act
                session.Refresh();

                // Assert
                Assert.True(File.Exists(tempPath));

                using var fs = File.OpenRead(tempPath);
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheet("Sheet1");
                Assert.Equal("FluentNPOI Works!", sheet.GetRow(0)?.GetCell(0)?.StringCellValue);
                Assert.Equal(123.0, sheet.GetRow(0)?.GetCell(1)?.NumericCellValue);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentHotReloadSession_RefreshCount_ShouldIncrement()
        {
            // Arrange
            var tempPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

            try
            {
                using var session = new FluentHotReloadSession(tempPath, wb =>
                {
                    wb.UseSheet("Sheet1").SetCellPosition(ExcelCol.A, 1).SetValue("Test");
                });

                // Act
                session.Refresh();
                session.Refresh();
                session.Refresh();

                // Assert
                Assert.Equal(3, session.RefreshCount);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [Fact]
        public void FluentLivePreview_CreateSession_ShouldReturnSession()
        {
            // Arrange & Act
            using var session = FluentLivePreview.CreateSession("test.xlsx", wb =>
            {
                wb.UseSheet("Sheet1").SetCellPosition(ExcelCol.A, 1).SetValue("Test");
            });

            // Assert
            Assert.NotNull(session);
            Assert.Equal("test.xlsx", session.OutputPath);
        }

        #endregion
    }
}
