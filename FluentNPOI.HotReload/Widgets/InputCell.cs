using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;
using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget representing an editable input cell with a default value.
/// Used for cells where users can enter data. Supports state preservation
/// during hot reload when a key is provided.
/// </summary>
public class InputCell : ExcelWidget
{
    /// <summary>
    /// Gets the default value for this cell.
    /// </summary>
    public object? DefaultValue { get; }

    /// <summary>
    /// Gets or sets the background color for the input cell.
    /// </summary>
    public IndexedColors? Background { get; set; }

    /// <summary>
    /// Gets or sets whether to add a visual indicator that this is an editable cell.
    /// </summary>
    public bool ShowEditableIndicator { get; set; } = true;

    /// <summary>
    /// Creates a new InputCell widget with the specified default value.
    /// </summary>
    /// <param name="defaultValue">The default value for the cell.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public InputCell(
        object? defaultValue = null,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        DefaultValue = defaultValue;
    }

    /// <summary>
    /// Sets the background color and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="color">The background color.</param>
    /// <returns>This InputCell instance.</returns>
    public InputCell SetBackground(IndexedColors color)
    {
        Background = color;
        return this;
    }

    /// <summary>
    /// Sets whether to show an editable indicator and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="show">Whether to show the indicator.</param>
    /// <returns>This InputCell instance.</returns>
    public InputCell SetShowEditableIndicator(bool show)
    {
        ShowEditableIndicator = show;
        return this;
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        ctx.SetValue(DefaultValue);

        if (Background != null)
        {
            ctx.SetBackgroundColor(Background);
        }
        else if (ShowEditableIndicator)
        {
            // Light yellow background to indicate editable cell
            ctx.SetBackgroundColor(IndexedColors.LightYellow);
        }

        // Add debug comment with source location if key is set
        if (!string.IsNullOrEmpty(Key))
        {
            ctx.SetComment($"Key: {Key}\n{SourceLocation.ToComment()}");
        }
    }
}
