using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget that arranges its children horizontally (one per column).
/// Each child is rendered in the next column, advancing rightward.
/// </summary>
public class Row : ExcelWidget
{
    /// <summary>
    /// Gets the list of child widgets (cells).
    /// </summary>
    public List<ExcelWidget> Cells { get; }

    /// <summary>
    /// Creates a new Row widget with the specified cells.
    /// </summary>
    /// <param name="cells">The child widgets to arrange horizontally.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public Row(
        IEnumerable<ExcelWidget> cells,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        Cells = cells.ToList();
    }

    /// <summary>
    /// Creates a new Row widget with the specified cells.
    /// </summary>
    /// <param name="cells">The child widgets to arrange horizontally.</param>
    public Row(params ExcelWidget[] cells)
        : this(cells.AsEnumerable())
    {
    }

    /// <summary>
    /// Creates a new Row widget with a key and specified cells.
    /// </summary>
    /// <param name="key">The key for this row.</param>
    /// <param name="cells">The child widgets to arrange horizontally.</param>
    public Row(string key, IEnumerable<ExcelWidget> cells)
        : this(cells)
    {
        Key = key;
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        foreach (var cell in Cells)
        {
            cell.Build(ctx);
            ctx.MoveToNextColumn();
        }
    }
}
