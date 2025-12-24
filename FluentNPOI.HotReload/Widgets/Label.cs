using System.Runtime.CompilerServices;
using FluentNPOI.HotReload.Context;
using NPOI.SS.UserModel;

namespace FluentNPOI.HotReload.Widgets;

/// <summary>
/// A widget representing a simple text label cell.
/// Used for displaying static text content.
/// </summary>
public class Label : ExcelWidget
{
    /// <summary>
    /// Gets the label text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets or sets whether the text is bold.
    /// </summary>
    public bool IsBold { get; set; }

    /// <summary>
    /// Creates a new Label widget with the specified text.
    /// </summary>
    /// <param name="text">The label text.</param>
    /// <param name="filePath">Auto-captured source file path.</param>
    /// <param name="lineNumber">Auto-captured line number.</param>
    public Label(
        string text,
        [CallerFilePath] string filePath = "",
        [CallerLineNumber] int lineNumber = 0)
        : base(filePath, lineNumber)
    {
        Text = text;
    }

    /// <summary>
    /// Sets whether the text is bold and returns this instance for fluent chaining.
    /// </summary>
    /// <param name="bold">Whether to make the text bold.</param>
    /// <returns>This Label instance.</returns>
    public Label SetBold(bool bold = true)
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
    }
}
