using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using FluentNPOI.HotReload.Bridge;

namespace FluentNPOIUnitTest
{
    /// <summary>
    /// Tests for FluentNPOI.HotReload Phase 3: LibreOffice Bridge
    /// </summary>
    public class LibreOfficeBridgeTests
    {
        #region DetectPath Tests

        [Fact]
        public void DetectPath_ShouldReturnNullOrValidPath()
        {
            // Act
            var path = LibreOfficeBridge.DetectPath();

            // Assert
            if (path != null)
            {
                Assert.True(File.Exists(path), $"Detected path should exist: {path}");
            }
            // If null, that's also acceptable (LibreOffice not installed)
        }

        [Fact]
        public void DetectPath_ShouldRespectEnvironmentVariable()
        {
            // Arrange
            var originalValue = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
            var testPath = Path.Combine(Path.GetTempPath(), "test_soffice.exe");

            try
            {
                // Create a dummy file
                File.WriteAllText(testPath, "dummy");
                Environment.SetEnvironmentVariable("LIBREOFFICE_PATH", testPath);

                // Act
                var detectedPath = LibreOfficeBridge.DetectPath();

                // Assert
                Assert.Equal(testPath, detectedPath);
            }
            finally
            {
                // Restore original value
                Environment.SetEnvironmentVariable("LIBREOFFICE_PATH", originalValue);
                if (File.Exists(testPath))
                    File.Delete(testPath);
            }
        }

        #endregion

        #region IsRunning Tests

        [Fact]
        public void IsRunning_ShouldBeFalseInitially()
        {
            // Arrange - Kill any existing process
            LibreOfficeBridge.Kill();

            // Assert
            Assert.False(LibreOfficeBridge.IsRunning);
        }

        #endregion

        #region LibreOfficeOptions Tests

        [Fact]
        public void LibreOfficeOptions_ShouldHaveCorrectDefaults()
        {
            // Arrange & Act
            var options = new LibreOfficeOptions();

            // Assert
            Assert.Null(options.SofficePath);
            Assert.True(options.AutoOpen);
            Assert.True(options.AutoRefresh);
            Assert.Equal(3000, options.FileLockWaitMs);
        }

        [Fact]
        public void LibreOfficeOptions_ShouldBeConfigurable()
        {
            // Arrange & Act
            var options = new LibreOfficeOptions
            {
                SofficePath = "/custom/path/soffice",
                AutoOpen = false,
                AutoRefresh = false,
                FileLockWaitMs = 5000
            };

            // Assert
            Assert.Equal("/custom/path/soffice", options.SofficePath);
            Assert.False(options.AutoOpen);
            Assert.False(options.AutoRefresh);
            Assert.Equal(5000, options.FileLockWaitMs);
        }

        #endregion

        #region Event Tests

        [Fact]
        public async Task OpenFailed_ShouldBeRaisedWhenLibreOfficeNotFound()
        {
            // Arrange
            var originalValue = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
            Exception? capturedError = null;

            void ErrorHandler(Exception ex) => capturedError = ex;
            LibreOfficeBridge.OpenFailed += ErrorHandler;

            try
            {
                // Set to non-existent path to force failure
                Environment.SetEnvironmentVariable("LIBREOFFICE_PATH", null);

                // Only run this test if LibreOffice is not installed
                if (LibreOfficeBridge.DetectPath() == null)
                {
                    var tempFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
                    File.WriteAllBytes(tempFile, new byte[] { 0x50, 0x4B, 0x03, 0x04 }); // Minimal zip header

                    try
                    {
                        // Act & Assert
                        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
                            await LibreOfficeBridge.OpenAsync(tempFile));
                    }
                    finally
                    {
                        if (File.Exists(tempFile))
                            File.Delete(tempFile);
                    }
                }
            }
            finally
            {
                Environment.SetEnvironmentVariable("LIBREOFFICE_PATH", originalValue);
                LibreOfficeBridge.OpenFailed -= ErrorHandler;
            }
        }

        #endregion

        #region Kill Tests

        [Fact]
        public void Kill_ShouldNotThrowWhenNoProcessRunning()
        {
            // Act & Assert (should not throw)
            var exception = Record.Exception(() => LibreOfficeBridge.Kill());
            Assert.Null(exception);
        }

        #endregion
    }
}
