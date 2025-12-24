using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;
using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget representing a header cell with styled text.
/// Typically used for column headers with bold text and background color.
/// </summary>
public class Header : ExcelWidget
{
    /// <summary>
    /// Gets the header text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets or sets the background color.
    /// </summary>
    public IndexedColors? Background { get; set; }

    /// <summary>
    /// Gets or sets whether the text is bold. Default is true for headers.
    /// </summary>
    public bool IsBold { get; set; } = true;

    /// <summary>
    /// Creates a new Header widget with the specified text.
    /// </summary>
    /// <param name="text">The header text.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public Header(
        string text,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        Text = text;
    }

    /// <summary>
    /// Sets the background color and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="color">The background color.</param>
    /// <returns>This Header instance.</returns>
    public Header SetBackground(IndexedColors color)
    {
        Background = color;
        return this;
    }

    /// <summary>
    /// Sets whether the text is bold and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="bold">Whether to make the text bold.</param>
    /// <returns>This Header instance.</returns>
    public Header SetBold(bool bold)
    {
        IsBold = bold;
        return this;
    }

    /// <inheritdoc/>
    public override void Build(ExcelContext ctx)
    {
        ctx.SetValue(Text);

        if (IsBold)
        {
            ctx.SetBold();
        }

        if (Background != null)
        {
            ctx.SetBackgroundColor(Background);
        }
    }
}
