using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget that arranges its children vertically (one per row).
/// Each child is rendered on its own row, advancing downward.
/// </summary>
public class Column : ExcelWidget
{
    /// <summary>
    /// Gets the list of child widgets.
    /// </summary>
    public List<ExcelWidget> Children { get; }

    /// <summary>
    /// Creates a new Column widget with the specified children.
    /// </summary>
    /// <param name="children">The child widgets to arrange vertically.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public Column(
        IEnumerable<ExcelWidget> children,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        Children = children.ToList();
    }

    /// <summary>
    /// Creates a new Column widget with the specified children.
    /// </summary>
    /// <param name="children">The child widgets to arrange vertically.</param>
    public Column(params ExcelWidget[] children)
        : this(children.AsEnumerable())
    {
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        foreach (var child in Children)
        {
            child.Build(ctx);
            ctx.MoveToNextRow();
        }
    }
}
