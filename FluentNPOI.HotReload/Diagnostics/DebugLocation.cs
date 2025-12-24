using System.IO;

namespace FluentNPOI.HotReload.Diagnostics;

/// <summary>
/// Records the source location (file path and line number) where a widget is defined.
/// Used for debugging and source mapping during hot reload.
/// </summary>
/// <param name="FilePath">The source file path where the widget was created.</param>
/// <param name="LineNumber">The line number in the source file.</param>
public record DebugLocation(string FilePath, int LineNumber)
{
    /// <summary>
    /// Converts the debug location to a comment string for embedding in Excel cells.
    /// </summary>
    /// <returns>A formatted string showing the file name and line number.</returns>
    public string ToComment() => $"Source: {Path.GetFileName(FilePath)}:{LineNumber}";

    /// <summary>
    /// Returns the full path with line number for IDE navigation.
    /// </summary>
    public override string ToString() => $"{FilePath}:{LineNumber}";
}
