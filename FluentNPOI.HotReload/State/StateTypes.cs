namespace FluentNPOI.HotReload.State;

/// <summary>
/// Specifies where to persist widget state values.
/// </summary>
public enum StateStorageStrategy
{
    /// <summary>
    /// Store state in cell comments. Most visible to users.
    /// </summary>
    Comments,

    /// <summary>
    /// Store state in CustomXmlParts. Invisible to users but requires .xlsx format.
    /// This is the recommended option for production use.
    /// </summary>
    CustomXmlParts,

    /// <summary>
    /// Store state in a hidden sheet. Compatible with all Excel versions.
    /// </summary>
    HiddenSheet,

    /// <summary>
    /// Do not persist state. Values are lost on each refresh.
    /// </summary>
    None
}

/// <summary>
/// Represents a cell location for state tracking.
/// </summary>
public record CellLocation(string SheetName, int Row, int Column)
{
    public override string ToString() => $"{SheetName}!{(char)('A' + Column)}{Row}";
}

/// <summary>
/// Represents a persisted state entry.
/// </summary>
public class StateEntry
{
    /// <summary>
    /// The unique key identifying this state.
    /// </summary>
    public string Key { get; set; } = string.Empty;

    /// <summary>
    /// The stored value.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// The cell location where this value came from.
    /// </summary>
    public CellLocation? Location { get; set; }

    /// <summary>
    /// The timestamp when this value was last updated.
    /// </summary>
    public DateTime LastUpdated { get; set; } = DateTime.Now;
}
