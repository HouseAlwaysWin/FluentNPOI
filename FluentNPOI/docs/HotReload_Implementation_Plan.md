# NPOI Fluent API Declarative Hot Reload System

Building a declarative UI-style Excel development experience with live preview capabilities through .NET Hot Reload integration.

## User Review Required

> [!IMPORTANT]
> **New Project Required**: This feature requires creating a new `FluentNPOI.HotReload` project to keep the core library lightweight.

> [!WARNING]
> **LibreOffice Dependency**: The live preview feature requires LibreOffice Calc to be installed on the developer's machine for the intended workflow.

> [!CAUTION]
> **Hot Reload Limitations**: .NET Hot Reload has restrictions (no new fields in structs, limited lambda changes). Some code changes will require full rebuild.

---

## Proposed Changes

### FluentNPOI.HotReload (NEW Project)

Create a new project to house all hot reload and widget infrastructure.

#### [NEW] FluentNPOI.HotReload.csproj
- Target: `net8.0`
- Dependencies: `FluentNPOI`, `Microsoft.VisualStudio.HotReload.Components`
- Package as optional NuGet extension

---

### Widget Engine Core

#### [NEW] Widgets/ExcelWidget.cs

Base class for all declarative Excel widgets:

```csharp
public abstract class ExcelWidget
{
    public string? Key { get; set; }
    public DebugLocation SourceLocation { get; }
    
    protected ExcelWidget(
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
    {
        SourceLocation = new DebugLocation(filePath, lineNumber);
    }
    
    public abstract void Build(ExcelContext ctx);
}
```

#### [NEW] Widgets/Column.cs, Row.cs, Header.cs, Label.cs, InputCell.cs

Core widget implementations following Flutter-style composition:

```csharp
public class Column : ExcelWidget
{
    public List<ExcelWidget> Children { get; }
    
    public Column(IEnumerable<ExcelWidget> children) => Children = children.ToList();
    
    public override void Build(ExcelContext ctx)
    {
        foreach (var child in Children)
        {
            child.Build(ctx);
            ctx.MoveToNextRow();
        }
    }
}

public class Header : ExcelWidget
{
    public string Text { get; }
    public IndexedColors? Background { get; set; }
    
    public Header(string text) => Text = text;
    
    public Header SetBackground(IndexedColors color)
    {
        Background = color;
        return this;
    }
    
    public override void Build(ExcelContext ctx)
    {
        ctx.SetValue(Text);
        if (Background.HasValue)
            ctx.SetBackgroundColor(Background.Value);
    }
}
```

#### [NEW] Context/ExcelContext.cs

Virtual sheet builder that accumulates operations before flushing to NPOI:

```csharp
public class ExcelContext
{
    private readonly FluentSheet _sheet;
    private int _currentRow = 1;
    private ExcelCol _currentCol = ExcelCol.A;
    
    public void SetValue(object value) => 
        _sheet.SetCellPosition(_currentCol, _currentRow).SetValue(value);
    
    public void MoveToNextRow() => _currentRow++;
    public void MoveToNextColumn() => _currentCol++;
}
```

---

### Hot Reload Infrastructure

#### [NEW] HotReload/HotReloadSession.cs

Manages the hot reload lifecycle and triggers refresh:

```csharp
public class HotReloadSession
{
    private readonly string _outputPath;
    private readonly Func<ExcelWidget> _rootWidgetFactory;
    
    public void Start()
    {
        HotReloadHandler.RefreshRequested += OnRefreshRequested;
        Refresh(); // Initial build
        LibreOfficeBridge.Open(_outputPath);
    }
    
    private void OnRefreshRequested(Type[]? updatedTypes)
    {
        Refresh();
    }
    
    public void Refresh()
    {
        var widget = _rootWidgetFactory();
        var workbook = new FluentWorkbook(new XSSFWorkbook());
        var ctx = new ExcelContext(workbook.UseSheet("Sheet1", true));
        widget.Build(ctx);
        workbook.SaveToPath(_outputPath);
    }
}
```

#### [NEW] HotReload/HotReloadHandler.cs

MetadataUpdateHandler integration:

```csharp
[assembly: MetadataUpdateHandler(typeof(HotReloadHandler))]

internal static class HotReloadHandler
{
    public static event Action<Type[]?>? RefreshRequested;
    
    // Called by .NET runtime when hot reload occurs
    public static void UpdateApplication(Type[]? updatedTypes)
    {
        RefreshRequested?.Invoke(updatedTypes);
    }
}
```

---

### LibreOffice Bridge

#### [NEW] Bridge/LibreOfficeBridge.cs

Manages LibreOffice process lifecycle:

```csharp
public static class LibreOfficeBridge
{
    private static Process? _process;
    
    public static string? DetectPath()
    {
        var paths = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            Environment.GetEnvironmentVariable("LIBREOFFICE_PATH")
        };
        return paths.FirstOrDefault(File.Exists);
    }
    
    public static void Open(string filePath)
    {
        Kill(); // Close previous instance
        
        var shadowPath = CreateShadowCopy(filePath);
        var sofficePath = DetectPath() ?? throw new Exception("LibreOffice not found");
        
        _process = Process.Start(new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments = $"--nologo --calc \"{shadowPath}\"",
            UseShellExecute = false
        });
    }
    
    public static void Kill()
    {
        foreach (var proc in Process.GetProcessesByName("soffice.bin"))
            proc.Kill();
    }
    
    private static string CreateShadowCopy(string original)
    {
        var shadow = Path.Combine(Path.GetTempPath(), 
            $"FluentNPOI_Preview_{Path.GetFileName(original)}");
        File.Copy(original, shadow, overwrite: true);
        return shadow;
    }
}
```

---

### Keyed State Management

#### [NEW] State/KeyedStateManager.cs

Captures and restores user-modified values:

```csharp
public class KeyedStateManager
{
    private readonly Dictionary<string, object?> _capturedValues = new();
    
    public void Capture(FluentWorkbook workbook)
    {
        foreach (var sheet in workbook.GetSheetNames())
        {
            var fluentSheet = workbook.UseSheet(sheet);
            // Scan for cells with Key metadata in comments
            // Store values indexed by key
        }
    }
    
    public void Restore(ExcelContext ctx)
    {
        // During build, if a widget has a Key and we have a captured value,
        // restore it instead of using default
    }
}
```

#### [NEW] Debug/DebugLocation.cs

Source location tracking:

```csharp
public record DebugLocation(string FilePath, int LineNumber)
{
    public string ToComment() => $"Source: {Path.GetFileName(FilePath)}:{LineNumber}";
}
```

---

### Integration

#### [NEW] ExcelLivePreview.cs

High-level API for developers:

```csharp
public static class ExcelLivePreview
{
    public static void Run<TWidget>(string outputPath) where TWidget : ExcelWidget, new()
    {
        var session = new HotReloadSession(outputPath, () => new TWidget());
        session.Start();
        
        Console.WriteLine("Hot Reload active. Press Ctrl+C to exit.");
        Console.WriteLine($"Preview: {outputPath}");
        
        // Keep app alive
        Thread.Sleep(Timeout.Infinite);
    }
}
```

---

## Verification Plan

### Automated Tests

**1. Widget Rendering Tests**
```bash
dotnet test FluentNPOIUnitTest --filter "FullyQualifiedName~WidgetRendering"
```

**2. Context Tests**
```bash
dotnet test FluentNPOIUnitTest --filter "FullyQualifiedName~ExcelContext"
```

**3. LibreOffice Bridge Tests**
```bash
dotnet test FluentNPOIUnitTest --filter "FullyQualifiedName~LibreOfficeBridge"
```

### Manual Verification

1. Create example console app with `SalesReport : ExcelWidget`
2. Run with `dotnet watch run`
3. Modify style in code and save
4. Verify LibreOffice refreshes within 3 seconds

---

## Project Structure

```
FluentNPOI/
├── FluentNPOI/                    # Existing core library
├── FluentNPOI.HotReload/          # NEW: Hot reload extension
│   ├── FluentNPOI.HotReload.csproj
│   ├── Widgets/
│   │   ├── ExcelWidget.cs
│   │   ├── Column.cs
│   │   ├── Row.cs
│   │   ├── Header.cs
│   │   ├── Label.cs
│   │   └── InputCell.cs
│   ├── Context/
│   │   └── ExcelContext.cs
│   ├── HotReload/
│   │   ├── HotReloadHandler.cs
│   │   └── HotReloadSession.cs
│   ├── Bridge/
│   │   └── LibreOfficeBridge.cs
│   ├── State/
│   │   └── KeyedStateManager.cs
│   ├── Debug/
│   │   └── DebugLocation.cs
│   └── ExcelLivePreview.cs
└── FluentNPOIUnitTest/
```

---

## API Preview

```csharp
// Program.cs
using FluentNPOI.HotReload;

ExcelLivePreview.Run<SalesReport>("output/preview.xlsx");

// SalesReport.cs
public class SalesReport : ExcelWidget
{
    public override void Build(ExcelContext ctx)
    {
        new Column(children: [
            new Header("年度銷售預算")
                .SetBackground(IndexedColors.RoyalBlue),
            
            new Row(key: "input_area", cells: [
                new Label("調整係數"),
                new InputCell(defaultValue: 1.2)
            ])
        ]).Build(ctx);
    }
}
```

---

## Timeline Estimate

| Phase | Estimated Time |
|-------|---------------|
| Phase 1: Widget Engine | 4-6 hours |
| Phase 2: Hot Reload Infra | 3-4 hours |
| Phase 3: LibreOffice Bridge | 2-3 hours |
| Phase 4: Keyed State | 3-4 hours |
| Phase 5: Testing & Docs | 2-3 hours |
| **Total** | **14-20 hours** |
