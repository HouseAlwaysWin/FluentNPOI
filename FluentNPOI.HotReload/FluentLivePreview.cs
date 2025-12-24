using FluentNPOI.HotReload.Bridge;
using FluentNPOI.HotReload.HotReload;
using FluentNPOI.Stages;
using NPOI.XSSF.UserModel;

namespace FluentNPOI.HotReload;

/// <summary>
/// Hot reload support for existing FluentNPOI code without using the widget system.
/// This allows gradual adoption - use your existing code with hot reload today,
/// migrate to widgets when ready.
/// </summary>
/// <example>
/// <code>
/// // Use your existing FluentNPOI code with hot reload
/// FluentLivePreview.Run("output.xlsx", workbook =>
/// {
///     workbook.UseSheet("Sheet1")
///         .SetCellPosition(ExcelCol.A, 1)
///         .SetValue("Hello World")
///         .SetCellPosition(ExcelCol.A, 2)
///         .SetValue("Hot Reload Works!");
/// });
/// </code>
/// </example>
public static class FluentLivePreview
{
    /// <summary>
    /// Runs a hot reload session with existing FluentNPOI code.
    /// No need to learn the widget system - just wrap your existing code.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="buildAction">Your existing FluentNPOI code.</param>
    /// <param name="configure">Optional session configuration.</param>
    public static void Run(
        string outputPath,
        Action<FluentWorkbook> buildAction,
        Action<FluentHotReloadSession>? configure = null)
    {
        using var session = new FluentHotReloadSession(outputPath, buildAction);
        configure?.Invoke(session);
        session.Start();

        PrintBanner(outputPath);

        // Set up graceful shutdown
        var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
            Console.WriteLine("\n🛑 Shutting down...");
        };

        try
        {
            Task.Delay(Timeout.Infinite, cts.Token).Wait();
        }
        catch (OperationCanceledException) { }
        catch (AggregateException ex) when (ex.InnerException is OperationCanceledException) { }

        Console.WriteLine("👋 Session ended. Goodbye!");
    }

    /// <summary>
    /// Runs a hot reload session asynchronously with existing FluentNPOI code.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="buildAction">Your existing FluentNPOI code.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <param name="configure">Optional session configuration.</param>
    public static async Task RunAsync(
        string outputPath,
        Action<FluentWorkbook> buildAction,
        CancellationToken cancellationToken = default,
        Action<FluentHotReloadSession>? configure = null)
    {
        using var session = new FluentHotReloadSession(outputPath, buildAction);
        configure?.Invoke(session);
        session.Start();

        PrintBanner(outputPath);

        try
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
        }
        catch (OperationCanceledException) { }

        Console.WriteLine("👋 Session ended.");
    }

    /// <summary>
    /// Creates a session for existing FluentNPOI code without starting it.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel preview file to.</param>
    /// <param name="buildAction">Your existing FluentNPOI code.</param>
    /// <returns>A configured but not started FluentHotReloadSession.</returns>
    public static FluentHotReloadSession CreateSession(
        string outputPath,
        Action<FluentWorkbook> buildAction)
    {
        return new FluentHotReloadSession(outputPath, buildAction);
    }

    private static void PrintBanner(string outputPath)
    {
        Console.WriteLine();
        Console.WriteLine("╔═══════════════════════════════════════════════════════════╗");
        Console.WriteLine("║        🔥 FluentNPOI Hot Reload Active 🔥                 ║");
        Console.WriteLine("╠═══════════════════════════════════════════════════════════╣");
        Console.WriteLine($"║  📊 Preview: {TruncatePath(outputPath, 44),-44} ║");
        Console.WriteLine("║                                                           ║");
        Console.WriteLine("║  • Edit your FluentNPOI code and save to see changes      ║");
        Console.WriteLine("║  • Press Ctrl+C to exit                                   ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════════╝");
        Console.WriteLine();
    }

    private static string TruncatePath(string path, int maxLength)
    {
        var fullPath = Path.GetFullPath(path);
        if (fullPath.Length <= maxLength)
            return fullPath;
        return "..." + fullPath.Substring(fullPath.Length - maxLength + 3);
    }
}

/// <summary>
/// Hot reload session for existing FluentNPOI code (imperative style).
/// </summary>
public class FluentHotReloadSession : IDisposable
{
    private readonly string _outputPath;
    private readonly Action<FluentWorkbook> _buildAction;
    private readonly object _refreshLock = new();
    private bool _isDisposed;
    private int _refreshCount;

    /// <summary>
    /// Event raised after a successful refresh.
    /// </summary>
    public event Action<int>? RefreshCompleted;

    /// <summary>
    /// Event raised when an error occurs during refresh.
    /// </summary>
    public event Action<Exception>? RefreshError;

    /// <summary>
    /// Gets the number of refreshes performed in this session.
    /// </summary>
    public int RefreshCount => _refreshCount;

    /// <summary>
    /// Gets the output file path.
    /// </summary>
    public string OutputPath => _outputPath;

    /// <summary>
    /// Gets or sets whether to automatically exit on Rude Edit detection.
    /// </summary>
    public bool AutoRestartOnRudeEdit { get; set; } = true;

    /// <summary>
    /// Gets or sets the LibreOffice options for auto-open and auto-refresh.
    /// </summary>
    public LibreOfficeOptions LibreOfficeOptions { get; set; } = new();

    // Shadow copy path for LibreOffice
    private string? _shadowCopyPath;
    private System.Diagnostics.Process? _libreOfficeProcess;

    /// <summary>
    /// Creates a new FluentHotReloadSession.
    /// </summary>
    /// <param name="outputPath">The path to write the Excel file to.</param>
    /// <param name="buildAction">The FluentNPOI building action.</param>
    public FluentHotReloadSession(string outputPath, Action<FluentWorkbook> buildAction)
    {
        _outputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
        _buildAction = buildAction ?? throw new ArgumentNullException(nameof(buildAction));

        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Starts the hot reload session.
    /// </summary>
    public void Start()
    {
        HotReloadHandler.RefreshRequested += OnRefreshRequested;
        HotReloadHandler.RudeEditDetected += OnRudeEditDetected;
        Refresh();

        Console.WriteLine($"🔥 Fluent Hot Reload session started");
        Console.WriteLine($"📊 Output: {Path.GetFullPath(_outputPath)}");

        // Open LibreOffice if configured
        if (LibreOfficeOptions.AutoOpen)
        {
            _ = TryOpenLibreOfficeAsync();
        }
    }

    /// <summary>
    /// Stops the hot reload session.
    /// </summary>
    public void Stop()
    {
        HotReloadHandler.RefreshRequested -= OnRefreshRequested;
        HotReloadHandler.RudeEditDetected -= OnRudeEditDetected;
        Console.WriteLine("👋 Fluent Hot Reload session stopped");
    }

    private void OnRefreshRequested(Type[]? updatedTypes)
    {
        // If not using shadow copy, we MUST kill LibreOffice before Refresh() attempts to write to the file
        if (!LibreOfficeOptions.UseShadowCopy && LibreOfficeOptions.AutoRefresh)
        {
            Console.WriteLine("🛑 Closing LibreOffice to release file lock...");
            KillExistingSofficeProcesses();
            // Wait for file handles to be strictly released
            Thread.Sleep(500);
        }

        Refresh();

        // Auto-refresh LibreOffice if configured
        if (LibreOfficeOptions.AutoRefresh)
        {
            _ = TryRefreshLibreOfficeAsync();
        }
    }

    private async Task TryOpenLibreOfficeAsync()
    {
        try
        {
            string targetPath;

            if (LibreOfficeOptions.UseShadowCopy)
            {
                // Create fixed shadow copy path (not random GUID)
                _shadowCopyPath = GetShadowCopyPath(_outputPath);
                targetPath = _shadowCopyPath;

                // Copy current output to shadow
                if (File.Exists(_outputPath))
                {
                    File.Copy(_outputPath, _shadowCopyPath, overwrite: true);
                }
            }
            else
            {
                targetPath = _outputPath;
            }

            // Detect LibreOffice
            var sofficePath = LibreOfficeBridge.DetectPath();
            if (string.IsNullOrEmpty(sofficePath))
            {
                Console.WriteLine("⚠️ LibreOffice not found. Please install LibreOffice.");
                return;
            }

            // Kill existing soffice processes (for clean restart on Rude Edit)
            KillExistingSofficeProcesses();
            await Task.Delay(300); // Wait for process to exit

            _libreOfficeProcess = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = $"--nologo --norestore --calc \"{targetPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            });

            Console.WriteLine($"📂 LibreOffice opened: {Path.GetFileName(targetPath)}");
            Console.WriteLine("   🔄 LibreOffice will automatically reopen after code changes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ LibreOffice open failed: {ex.Message}");
        }
    }

    private static void KillExistingSofficeProcesses()
    {
        try
        {
            foreach (var proc in System.Diagnostics.Process.GetProcessesByName("soffice"))
            {
                try { proc.Kill(); } catch { }
            }
            foreach (var proc in System.Diagnostics.Process.GetProcessesByName("soffice.bin"))
            {
                try { proc.Kill(); } catch { }
            }
        }
        catch { }
    }

    private async Task TryRefreshLibreOfficeAsync()
    {
        try
        {
            // Wait for NPOI to finish writing
            await Task.Delay(200);

            string targetPath;

            if (LibreOfficeOptions.UseShadowCopy)
            {
                // Update the shadow copy file
                if (!string.IsNullOrEmpty(_shadowCopyPath) && File.Exists(_outputPath))
                {
                    // Kill existing LibreOffice BEFORE copying to avoid file in use (if it's opening the shadow file)
                    if (_libreOfficeProcess != null && !_libreOfficeProcess.HasExited)
                    {
                        try { _libreOfficeProcess.Kill(); } catch { }
                    }

                    // Update shadow copy
                    File.Copy(_outputPath, _shadowCopyPath, overwrite: true);
                    targetPath = _shadowCopyPath;
                }
                else
                {
                    return; // Should not happen if initialized correctly
                }
            }
            else
            {
                targetPath = _outputPath;
                // Kill existing process
                if (_libreOfficeProcess != null && !_libreOfficeProcess.HasExited)
                {
                    try { _libreOfficeProcess.Kill(); } catch { }
                }
            }

            // Reopen LibreOffice
            var sofficePath = LibreOfficeBridge.DetectPath();
            if (!string.IsNullOrEmpty(sofficePath))
            {
                _libreOfficeProcess = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sofficePath,
                    Arguments = $"--nologo --norestore --calc \"{targetPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
                Console.WriteLine("🔄 LibreOffice reopened with updated file");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ LibreOffice refresh failed: {ex.Message}");
        }
    }

    private static string GetShadowCopyPath(string sourcePath)
    {
        var dir = Path.GetDirectoryName(sourcePath) ?? ".";
        var name = Path.GetFileNameWithoutExtension(sourcePath);
        var ext = Path.GetExtension(sourcePath);
        return Path.Combine(dir, $"{name}_preview{ext}");
    }

    private static string CreateShadowCopy(string sourcePath)
    {
        var shadowPath = GetShadowCopyPath(sourcePath);
        if (File.Exists(sourcePath))
        {
            File.Copy(sourcePath, shadowPath, overwrite: true);
        }
        return shadowPath;
    }

    private void OnRudeEditDetected()
    {
        Console.WriteLine("⚠️ Rude Edit detected");
        if (AutoRestartOnRudeEdit)
        {
            Console.WriteLine("🔄 Triggering full restart...");
            Environment.Exit(42);
        }
    }

    /// <summary>
    /// Performs a refresh, rebuilding and saving the Excel file.
    /// </summary>
    public void Refresh()
    {
        lock (_refreshLock)
        {
            try
            {
                var startTime = DateTime.Now;

                // Create new workbook
                var workbook = new XSSFWorkbook();
                var fluentWorkbook = new FluentWorkbook(workbook);

                // Execute user's FluentNPOI code
                _buildAction(fluentWorkbook);

                // Save to file - with fallback for locked files
                var savedPath = SaveWithFallback(fluentWorkbook, _outputPath);

                _refreshCount++;
                var elapsed = (DateTime.Now - startTime).TotalMilliseconds;

                Console.WriteLine($"✅ Refresh #{_refreshCount} completed in {elapsed:F0}ms");
                if (savedPath != _outputPath)
                {
                    Console.WriteLine($"   📁 File locked, saved to: {Path.GetFileName(savedPath)}");
                }
                RefreshCompleted?.Invoke(_refreshCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Refresh failed: {ex.Message}");
                RefreshError?.Invoke(ex);
            }
        }
    }

    /// <summary>
    /// Saves the workbook, using a fallback path if the file is locked.
    /// </summary>
    private string SaveWithFallback(FluentWorkbook workbook, string primaryPath)
    {
        // Retry logic to handle race conditions where file lock isn't released immediately
        int maxRetries = 10;
        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                workbook.SaveToPath(primaryPath);
                return primaryPath;
            }
            catch (IOException)
            {
                if (i < maxRetries - 1)
                {
                    Thread.Sleep(200); // Wait 200ms and retry
                }
            }
        }

        // File is definitely locked, use fallback with timestamp
        var dir = Path.GetDirectoryName(primaryPath) ?? ".";
        var name = Path.GetFileNameWithoutExtension(primaryPath);
        var ext = Path.GetExtension(primaryPath);
        var fallbackPath = Path.Combine(dir, $"{name}_{DateTime.Now:HHmmss}{ext}");

        try
        {
            workbook.SaveToPath(fallbackPath);
            Console.WriteLine($"⚠️ Target locked, saved to fallback: {Path.GetFileName(fallbackPath)}");
            return fallbackPath;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Failed to save even to fallback: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Disposes the session.
    /// </summary>
    public void Dispose()
    {
        if (!_isDisposed)
        {
            Stop();
            _isDisposed = true;
        }
    }
}

/// <summary>
/// Extension methods to add hot reload capability to existing FluentNPOI code.
/// </summary>
public static class FluentHotReloadExtensions
{
    /// <summary>
    /// Runs the FluentWorkbook builder with hot reload support.
    /// </summary>
    /// <param name="buildAction">Your FluentWorkbook building code.</param>
    /// <param name="outputPath">The output path for the preview file.</param>
    /// <param name="configure">Optional session configuration.</param>
    /// <example>
    /// <code>
    /// // One-liner to add hot reload to existing code
    /// Action&lt;FluentWorkbook&gt; myReport = wb =>
    /// {
    ///     wb.UseSheet("Data")
    ///         .SetCellPosition(ExcelCol.A, 1)
    ///         .SetValue("Report Data");
    /// };
    /// 
    /// myReport.WithHotReload("preview.xlsx");
    /// </code>
    /// </example>
    public static void WithHotReload(
        this Action<FluentWorkbook> buildAction,
        string outputPath,
        Action<FluentHotReloadSession>? configure = null)
    {
        FluentLivePreview.Run(outputPath, buildAction, configure);
    }
}
