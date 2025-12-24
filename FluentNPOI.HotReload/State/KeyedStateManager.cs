using System.Text.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FluentNPOI.HotReload.State;

/// <summary>
/// Manages keyed state persistence for widgets.
/// Preserves user-entered values across hot reload cycles.
/// </summary>
public class KeyedStateManager
{
    private readonly Dictionary<string, StateEntry> _stateCache = new();
    private const string StateSheetName = "__FluentNPOI_State__";
    private const string CustomXmlNamespace = "urn:fluentnpoi:hotreload:state";

    /// <summary>
    /// Gets or sets the storage strategy for state persistence.
    /// </summary>
    public StateStorageStrategy StorageStrategy { get; set; } = StateStorageStrategy.Comments;

    /// <summary>
    /// Gets the number of cached state entries.
    /// </summary>
    public int CachedEntryCount => _stateCache.Count;

    /// <summary>
    /// Clears all cached state.
    /// </summary>
    public void Clear()
    {
        _stateCache.Clear();
    }

    /// <summary>
    /// Sets a value for the given key.
    /// </summary>
    /// <param name="key">The unique key.</param>
    /// <param name="value">The value to store.</param>
    /// <param name="location">Optional cell location for tracking.</param>
    public void SetValue(string key, object? value, CellLocation? location = null)
    {
        if (string.IsNullOrEmpty(key))
            return;

        _stateCache[key] = new StateEntry
        {
            Key = key,
            Value = value,
            Location = location,
            LastUpdated = DateTime.Now
        };
    }

    /// <summary>
    /// Gets a value for the given key.
    /// </summary>
    /// <typeparam name="T">The expected type.</typeparam>
    /// <param name="key">The unique key.</param>
    /// <param name="defaultValue">Default value if key not found.</param>
    /// <returns>The stored value or default.</returns>
    public T? GetValue<T>(string key, T? defaultValue = default)
    {
        if (string.IsNullOrEmpty(key) || !_stateCache.TryGetValue(key, out var entry))
            return defaultValue;

        if (entry.Value is T typedValue)
            return typedValue;

        // Try conversion
        try
        {
            if (entry.Value != null)
                return (T)Convert.ChangeType(entry.Value, typeof(T));
        }
        catch
        {
            // Conversion failed
        }

        return defaultValue;
    }

    /// <summary>
    /// Checks if a key exists in the cache.
    /// </summary>
    public bool HasKey(string key)
    {
        return !string.IsNullOrEmpty(key) && _stateCache.ContainsKey(key);
    }

    /// <summary>
    /// Gets all cached state entries.
    /// </summary>
    public IEnumerable<StateEntry> GetAllEntries()
    {
        return _stateCache.Values;
    }

    /// <summary>
    /// Loads state from an existing workbook.
    /// </summary>
    /// <param name="workbook">The workbook to load state from.</param>
    public void LoadFromWorkbook(IWorkbook workbook)
    {
        if (StorageStrategy == StateStorageStrategy.None)
            return;

        switch (StorageStrategy)
        {
            case StateStorageStrategy.Comments:
                LoadFromComments(workbook);
                break;
            case StateStorageStrategy.HiddenSheet:
                LoadFromHiddenSheet(workbook);
                break;
            case StateStorageStrategy.CustomXmlParts:
                if (workbook is XSSFWorkbook xssfWorkbook)
                    LoadFromCustomXml(xssfWorkbook);
                else
                    LoadFromHiddenSheet(workbook); // Fallback
                break;
        }
    }

    /// <summary>
    /// Saves state to a workbook.
    /// </summary>
    /// <param name="workbook">The workbook to save state to.</param>
    public void SaveToWorkbook(IWorkbook workbook)
    {
        if (StorageStrategy == StateStorageStrategy.None || _stateCache.Count == 0)
            return;

        switch (StorageStrategy)
        {
            case StateStorageStrategy.Comments:
                // Comments are written inline during widget build
                break;
            case StateStorageStrategy.HiddenSheet:
                SaveToHiddenSheet(workbook);
                break;
            case StateStorageStrategy.CustomXmlParts:
                if (workbook is XSSFWorkbook xssfWorkbook)
                    SaveToCustomXml(xssfWorkbook);
                else
                    SaveToHiddenSheet(workbook); // Fallback
                break;
        }
    }

    /// <summary>
    /// Reads the current value from a cell and caches it.
    /// </summary>
    /// <param name="cell">The cell to read from.</param>
    /// <param name="key">The key to associate with this value.</param>
    public void CaptureValue(ICell? cell, string key)
    {
        if (cell == null || string.IsNullOrEmpty(key))
            return;

        object? value = cell.CellType switch
        {
            CellType.Numeric => cell.NumericCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Formula => cell.CellFormula,
            _ => null
        };

        var location = new CellLocation(
            cell.Sheet.SheetName,
            cell.RowIndex + 1,
            cell.ColumnIndex);

        SetValue(key, value, location);
    }

    #region Storage Implementations

    private void LoadFromComments(IWorkbook workbook)
    {
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            var sheet = workbook.GetSheetAt(i);
            if (sheet.SheetName == StateSheetName)
                continue;

            foreach (IRow row in sheet)
            {
                foreach (ICell cell in row)
                {
                    var comment = cell.CellComment?.String?.ToString();
                    if (string.IsNullOrEmpty(comment))
                        continue;

                    // Parse comment for key: Key: xxx
                    var lines = comment.Split('\n');
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("Key:"))
                        {
                            var key = line.Substring(4).Trim();
                            CaptureValue(cell, key);
                        }
                    }
                }
            }
        }
    }

    private void LoadFromHiddenSheet(IWorkbook workbook)
    {
        var stateSheet = workbook.GetSheet(StateSheetName);
        if (stateSheet == null)
            return;

        foreach (IRow row in stateSheet)
        {
            var keyCell = row.GetCell(0);
            var valueCell = row.GetCell(1);

            if (keyCell == null)
                continue;

            var key = keyCell.StringCellValue;
            object? value = valueCell?.CellType switch
            {
                CellType.Numeric => valueCell.NumericCellValue,
                CellType.String => valueCell.StringCellValue,
                CellType.Boolean => valueCell.BooleanCellValue,
                _ => null
            };

            SetValue(key, value);
        }
    }

    private void SaveToHiddenSheet(IWorkbook workbook)
    {
        // Remove existing state sheet
        var existingIndex = workbook.GetSheetIndex(StateSheetName);
        if (existingIndex >= 0)
            workbook.RemoveSheetAt(existingIndex);

        // Create new hidden state sheet
        var stateSheet = workbook.CreateSheet(StateSheetName);
        workbook.SetSheetHidden(workbook.GetSheetIndex(StateSheetName), SheetState.VeryHidden);

        int rowIndex = 0;
        foreach (var entry in _stateCache.Values)
        {
            var row = stateSheet.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue(entry.Key);

            var valueCell = row.CreateCell(1);
            switch (entry.Value)
            {
                case double d:
                    valueCell.SetCellValue(d);
                    break;
                case int i:
                    valueCell.SetCellValue(i);
                    break;
                case bool b:
                    valueCell.SetCellValue(b);
                    break;
                case string s:
                    valueCell.SetCellValue(s);
                    break;
                default:
                    valueCell.SetCellValue(entry.Value?.ToString() ?? "");
                    break;
            }
        }
    }

    private void LoadFromCustomXml(XSSFWorkbook workbook)
    {
        try
        {
            // Note: Full CustomXmlParts implementation requires POI 5.x special handling
            // For now, fallback to hidden sheet
            LoadFromHiddenSheet(workbook);
        }
        catch
        {
            // Silently fail
        }
    }

    private void SaveToCustomXml(XSSFWorkbook workbook)
    {
        try
        {
            // Note: Full CustomXmlParts implementation requires POI 5.x special handling
            // For now, fallback to hidden sheet
            SaveToHiddenSheet(workbook);
        }
        catch
        {
            // Silently fail
        }
    }

    #endregion

    /// <summary>
    /// Generates a comment string for a keyed widget.
    /// </summary>
    public static string GenerateKeyComment(string key, string? sourceInfo = null)
    {
        var comment = $"Key: {key}";
        if (!string.IsNullOrEmpty(sourceInfo))
            comment += $"\n{sourceInfo}";
        return comment;
    }
}
