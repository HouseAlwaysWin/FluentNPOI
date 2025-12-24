using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FluentNPOI.HotReload.Bridge;

/// <summary>
/// Manages LibreOffice Calc process for live Excel preview.
/// Creates shadow copies to avoid file locking issues.
/// </summary>
public static class LibreOfficeBridge
{
    private static Process? _currentProcess;
    private static string? _currentShadowPath;
    private static readonly object _lock = new();

    /// <summary>
    /// Event raised when LibreOffice is opened.
    /// </summary>
    public static event Action<string>? Opened;

    /// <summary>
    /// Event raised when LibreOffice fails to open.
    /// </summary>
    public static event Action<Exception>? OpenFailed;

    /// <summary>
    /// Gets whether LibreOffice is currently running.
    /// </summary>
    public static bool IsRunning => _currentProcess != null && !_currentProcess.HasExited;

    /// <summary>
    /// Detects the LibreOffice installation path based on the current OS.
    /// </summary>
    /// <returns>The path to soffice executable, or null if not found.</returns>
    public static string? DetectPath()
    {
        // Check environment variable first
        var envPath = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
        if (!string.IsNullOrEmpty(envPath) && File.Exists(envPath))
            return envPath;

        // OS-specific paths
        var paths = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            }
            : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? new[]
            {
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            }
            : new[]
            {
                "/usr/bin/libreoffice",
                "/usr/bin/soffice",
                "/usr/local/bin/libreoffice",
            };

        return paths.FirstOrDefault(File.Exists);
    }

    /// <summary>
    /// Opens an Excel file in LibreOffice Calc.
    /// Creates a shadow copy to avoid file locking issues during hot reload.
    /// </summary>
    /// <param name="filePath">The Excel file to open.</param>
    /// <param name="waitForLockMs">Maximum time to wait for file lock release.</param>
    public static async Task OpenAsync(string filePath, int waitForLockMs = 3000)
    {
        var sofficePath = DetectPath();
        if (string.IsNullOrEmpty(sofficePath))
        {
            var ex = new FileNotFoundException(
                "LibreOffice not found. Please install LibreOffice or set LIBREOFFICE_PATH environment variable.");
            OpenFailed?.Invoke(ex);
            throw ex;
        }

        try
        {
            // Wait for file to be unlocked
            await WaitForFileLockReleaseAsync(filePath, waitForLockMs);

            // Create shadow copy
            var shadowPath = CreateShadowCopy(filePath);

            // Kill existing process if running
            Kill();

            lock (_lock)
            {
                _currentShadowPath = shadowPath;
                _currentProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = sofficePath,
                    Arguments = $"--nologo --calc \"{shadowPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
            }

            Console.WriteLine($"📂 LibreOffice opened: {Path.GetFileName(filePath)}");
            Opened?.Invoke(shadowPath);
        }
        catch (Exception ex) when (ex is not FileNotFoundException)
        {
            OpenFailed?.Invoke(ex);
            throw;
        }
    }

    /// <summary>
    /// Opens an Excel file in LibreOffice Calc (synchronous version).
    /// </summary>
    /// <param name="filePath">The Excel file to open.</param>
    public static void Open(string filePath)
    {
        OpenAsync(filePath).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Refreshes LibreOffice by reopening with the latest version of the file.
    /// </summary>
    /// <param name="filePath">The Excel file to refresh.</param>
    public static async Task RefreshAsync(string filePath)
    {
        await OpenAsync(filePath);
    }

    /// <summary>
    /// Kills the currently running LibreOffice process.
    /// </summary>
    public static void Kill()
    {
        lock (_lock)
        {
            if (_currentProcess != null && !_currentProcess.HasExited)
            {
                try
                {
                    _currentProcess.Kill();
                    _currentProcess.WaitForExit(3000);
                    Console.WriteLine("🛑 LibreOffice closed");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Failed to close LibreOffice: {ex.Message}");
                }
            }
            _currentProcess = null;

            // Clean up shadow file
            CleanupShadowCopy();
        }
    }

    /// <summary>
    /// Waits for a file lock to be released.
    /// </summary>
    private static async Task WaitForFileLockReleaseAsync(string filePath, int maxWaitMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < maxWaitMs)
        {
            try
            {
                using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                return; // File is unlocked
            }
            catch (IOException)
            {
                await Task.Delay(50);
            }
            catch (UnauthorizedAccessException)
            {
                await Task.Delay(50);
            }
        }
        // Proceed anyway, might still work
    }

    /// <summary>
    /// Creates a shadow copy of the file to avoid locking issues.
    /// </summary>
    private static string CreateShadowCopy(string filePath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "FluentNPOI_HotReload");
        Directory.CreateDirectory(tempDir);

        var fileName = Path.GetFileName(filePath);
        var shadowPath = Path.Combine(tempDir, $"preview_{Guid.NewGuid():N}_{fileName}");

        File.Copy(filePath, shadowPath, overwrite: true);
        return shadowPath;
    }

    /// <summary>
    /// Cleans up the shadow copy file.
    /// </summary>
    private static void CleanupShadowCopy()
    {
        if (!string.IsNullOrEmpty(_currentShadowPath) && File.Exists(_currentShadowPath))
        {
            try
            {
                File.Delete(_currentShadowPath);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
        _currentShadowPath = null;
    }
}

/// <summary>
/// Options for configuring the LibreOffice bridge.
/// </summary>
public class LibreOfficeOptions
{
    /// <summary>
    /// Path to the LibreOffice soffice executable.
    /// If null, auto-detection will be used.
    /// </summary>
    public string? SofficePath { get; set; }

    /// <summary>
    /// Whether to auto-open LibreOffice on session start.
    /// Default is true.
    /// </summary>
    public bool AutoOpen { get; set; } = true;

    /// <summary>
    /// Whether to auto-refresh LibreOffice on each hot reload.
    /// Default is true.
    /// </summary>
    public bool AutoRefresh { get; set; } = true;

    /// <summary>
    /// Maximum time to wait for file lock release in milliseconds.
    /// Default is 3000ms.
    /// </summary>
    public int FileLockWaitMs { get; set; } = 3000;
}
