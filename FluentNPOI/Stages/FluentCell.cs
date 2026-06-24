using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Helpers;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// Cell operation class
    /// </summary>
    public class FluentCell : FluentCellBase
    {
        private ICell? _cell;
        private ExcelCol _col;
        private int _row;
        public FluentCell(IWorkbook workbook, ISheet? sheet,
        ICell cell, Dictionary<string, ICellStyle>? cellStylesCached = null)
            : base(workbook, sheet, cellStylesCached ?? new Dictionary<string, ICellStyle>())
        {
            _cell = cell;
            _col = (ExcelCol)cell.ColumnIndex;
            _row = cell.RowIndex;
        }

        public FluentCell SetValue<T>(T value)
        {
            if (_cell == null) return this;
            SetCellValue(_cell, value);
            return this;
        }

        public FluentCell SetFormulaValue(object value)
        {
            if (_cell == null) return this;
            SetFormulaValue(_cell, value);
            return this;
        }

        public FluentCell SetCellStyle(string cellStyleKey)
        {
            if (_cell == null) return this;

            if (!string.IsNullOrWhiteSpace(cellStyleKey) && _cellStylesCached.ContainsKey(cellStyleKey))
            {
                _cell.CellStyle = _cellStylesCached[cellStyleKey];
            }
            return this;
        }


        public FluentCell SetCellType(CellType cellType)
        {
            if (_cell == null) return this;
            _cell.SetCellType(cellType);
            return this;
        }

        /// <summary>
        /// Get current cell value
        /// </summary>
        /// <returns>Cell value (return bool, DateTime, double, string or null based on type)</returns>
        public object? GetValue()
        {
            return GetCellValue(_cell);
        }

        /// <summary>
        /// Get current cell value and convert to specified type
        /// </summary>
        /// <typeparam name="T">Target type</typeparam>
        /// <returns>Converted value</returns>
        public T GetValue<T>()
        {
            return GetCellValue<T>(_cell);
        }

        /// <summary>
        /// Get current cell formula string (if it is a formula cell)
        /// </summary>
        /// <returns>Formula string (without '=' prefix), or null if not a formula</returns>
        public string? GetFormula()
        {
            return GetCellFormulaValue(_cell);
        }

        /// <summary>
        /// Get current cell object
        /// </summary>
        /// <returns>NPOI ICell object</returns>
        public ICell? GetCell()
        {
            return _cell;
        }

        /// <summary>
        /// Set picture in cell (auto calculate height, keep aspect ratio)
        /// </summary>
        /// <param name="pictureBytes">Picture byte array</param>
        /// <param name="imgWidth">Image width (pixels)</param>
        /// <param name="anchorType">Anchor type</param>
        /// <param name="columnWidthRatio">Column width ratio (default 7.0, means pixel width divided by this value gets Excel column width in characters)</param>
        /// <returns>FluentCell instance, supports method chaining</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, AnchorType anchorType = AnchorType.MoveAndResize, double columnWidthRatio = 7.0)
        {
            // Auto calculate height (need to read original size from image)
            // Since cannot get image size directly from byte array, use width as height here (1:1 ratio)
            // If more precise ratio is needed, consider using System.Drawing.Image or other image libraries
            return SetPictureOnCell(pictureBytes, imgWidth, imgWidth, anchorType, columnWidthRatio);
        }

        /// <summary>
        /// Set picture in cell (manually set width and height)
        /// </summary>
        /// <param name="pictureBytes">Picture byte array</param>
        /// <param name="imgWidth">Image width (pixels)</param>
        /// <param name="imgHeight">Image height (pixels)</param>
        /// <param name="anchorType">Anchor type</param>
        /// <param name="columnWidthRatio">Column width ratio (default 7.0, means pixel width divided by this value gets Excel column width in characters)</param>
        /// <param name="pictureAction">Picture action delegate</param>
        /// <returns>FluentCell instance, supports method chaining</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, int imgHeight, AnchorType anchorType = AnchorType.MoveAndResize,
        double columnWidthRatio = 7.0, Action<IPicture>? pictureAction = null)
        {
            // Parameter validation
            ValidatePictureParameters(pictureBytes, imgWidth, imgHeight, columnWidthRatio);

            // Set column width
            double columnWidth = CalculateColumnWidth(imgWidth, columnWidthRatio);
            Sheet.SetColumnWidth((int)_col, (int)Math.Round(columnWidth * 256));

            // Get picture type and add to workbook
            var picType = GetPictureType(pictureBytes);
            int picIndex = _workbook.AddPicture(pictureBytes, picType);

            // Create drawing patriarch and anchor
            // NPOI 2.8.0 made IDrawing generic (IDrawing<T>); use var to bind the concrete type.
            var drawing = Sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = CreatePictureAnchor(imgWidth, imgHeight, anchorType);

            // Create picture
            IPicture pict = drawing.CreatePicture(anchor, picIndex);

            pictureAction?.Invoke(pict);

            return this;
        }

        /// <summary>
        /// Validate picture parameters
        /// </summary>
        private void ValidatePictureParameters(byte[] pictureBytes, int imgWidth, int imgHeight, double columnWidthRatio)
        {
            if (_cell == null)
            {
                throw new InvalidOperationException("No active cell. Call SetCellPosition(...) first.");
            }

            if (pictureBytes == null || pictureBytes.Length == 0)
            {
                throw new ArgumentException("Picture bytes cannot be null or empty.", nameof(pictureBytes));
            }

            if (imgWidth <= 0)
            {
                throw new ArgumentException("Image width must be greater than zero.", nameof(imgWidth));
            }

            if (imgHeight <= 0)
            {
                throw new ArgumentException("Image height must be greater than zero.", nameof(imgHeight));
            }

            if (columnWidthRatio <= 0)
            {
                throw new ArgumentException("Column width ratio must be greater than zero.", nameof(columnWidthRatio));
            }
        }

        /// <summary>
        /// Calculate column width (convert pixel width to Excel column width unit)
        /// </summary>
        /// <param name="imgWidth">Image width (pixels)</param>
        /// <param name="columnWidthRatio">Conversion ratio</param>
        /// <returns>Excel column width (characters)</returns>
        private double CalculateColumnWidth(int imgWidth, double columnWidthRatio)
        {
            // Excel column width unit: 1 character width = 256 units
            // Divide pixel width by conversion ratio to get character count
            return imgWidth / columnWidthRatio;
        }

        /// <summary>
        /// Create picture anchor, set complete position and size information
        /// </summary>
        /// <param name="imgWidth">Image width (pixels)</param>
        /// <param name="imgHeight">Image height (pixels)</param>
        /// <param name="anchorType">Anchor type</param>
        /// <returns>Configured IClientAnchor object</returns>
        private IClientAnchor CreatePictureAnchor(int imgWidth, int imgHeight, AnchorType anchorType)
        {
            ICreationHelper creationHelper = _workbook.GetCreationHelper();
            IClientAnchor anchor = creationHelper.CreateClientAnchor();

            // Set start position (_row is already 0-based, because converted in SetCellPosition)
            anchor.Col1 = (short)_col;
            anchor.Row1 = (short)_row;

            // Calculate end position (Col2 and Row2)
            // Calculate how many columns and rows needed based on image size and cell size
            // Excel default column width is about 8.43 characters (about 64 pixels), row height is about 15 pixels
            // Use simplified calculation here

            // Get current column width (in characters)
            // GetColumnWidth returns int (in 1/256 characters), convert to character count
            double columnWidthInChars = Sheet.GetColumnWidth((int)_col) / 256.0;

            // Get current row height (in points, 1 point ≈ 1.33 pixels)
            IRow row = Sheet.GetRow(_row) ?? Sheet.CreateRow(_row);
            short rowHeightInPoints = row.Height > 0 ? (short)(row.Height / 20.0) : (short)15; // Default row height about 15 points

            // Calculate columns needed (considering column width)
            // Assume 1 character width ≈ 7 pixels (adjust as needed)
            double pixelsPerChar = 7.0;
            double columnsNeeded = imgWidth / (columnWidthInChars * pixelsPerChar);
            short col2 = (short)Math.Min((int)_col + (int)Math.Ceiling(columnsNeeded), 16383); // Excel max column limit

            // Calculate rows needed (considering row height)
            // 1 點 ≈ 1.33 像素
            double pixelsPerPoint = 1.33;
            double rowsNeeded = imgHeight / (rowHeightInPoints * pixelsPerPoint);
            short row2 = (short)Math.Min(_row + (int)Math.Ceiling(rowsNeeded), 1048575); // Excel max row limit

            anchor.Col2 = col2;
            anchor.Row2 = row2;
            anchor.AnchorType = anchorType;

            return anchor;
        }

        private PictureType GetPictureType(byte[] pictureBytes)
        {
            if (pictureBytes == null || pictureBytes.Length < 4)
            {
                throw new ArgumentException("Invalid picture bytes: array is null or too short.", nameof(pictureBytes));
            }

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (pictureBytes.Length >= 8 &&
                pictureBytes[0] == 0x89 && pictureBytes[1] == 0x50 && pictureBytes[2] == 0x4E && pictureBytes[3] == 0x47 &&
                pictureBytes[4] == 0x0D && pictureBytes[5] == 0x0A && pictureBytes[6] == 0x1A && pictureBytes[7] == 0x0A)
            {
                return PictureType.PNG;
            }

            // JPEG: FF D8 FF
            if (pictureBytes.Length >= 3 &&
                pictureBytes[0] == 0xFF && pictureBytes[1] == 0xD8 && pictureBytes[2] == 0xFF)
            {
                return PictureType.JPEG;
            }

            // GIF: 47 49 46 38 (GIF8)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x47 && pictureBytes[1] == 0x49 && pictureBytes[2] == 0x46 && pictureBytes[3] == 0x38)
            {
                return PictureType.GIF;
            }

            // BMP/DIB: 42 4D (BM)
            if (pictureBytes.Length >= 2 &&
                pictureBytes[0] == 0x42 && pictureBytes[1] == 0x4D)
            {
                return PictureType.DIB;
            }

            // EMF: 01 00 00 00 (Check needs more checks, EMF files usually start with this)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x00 && pictureBytes[3] == 0x00)
            {
                // Check if valid EMF file (EMF header is usually 40 bytes)
                if (pictureBytes.Length >= 40)
                {
                    // Second DWORD of EMF file should be file size
                    // Simple check here, return EMF if matches EMF characteristics
                    return PictureType.EMF;
                }
            }

            // WMF: Usually starts with 01 00 09 00 (but need more checks)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x09 && pictureBytes[3] == 0x00)
            {
                return PictureType.WMF;
            }

            throw new NotSupportedException($"Unsupported picture format. File header: {BitConverter.ToString(pictureBytes, 0, Math.Min(8, pictureBytes.Length))}");
        }

        /// <summary>
        /// Set current operation cell position
        /// </summary>
        /// <param name="col">Column position</param>
        /// <param name="row">Row position (1-based)</param>
        public FluentCell SetCellPosition(ExcelCol col, int row)
        {
            _cell = SetCellPositionInternal(col, row);
            _col = col;
            _row = NormalizeRow(row);  // Store 0-based row
            return this;
        }

        /// <summary>
        /// Set formula (without '=' prefix)
        /// </summary>
        /// <param name="formula">Formula string (e.g. "SUM(A1:A10)")</param>
        public FluentCell SetFormula(string formula)
        {
            if (_cell == null) return this;
            if (string.IsNullOrWhiteSpace(formula)) return this;

            // Remove '=' prefix (if exists)
            if (formula.StartsWith("=")) formula = formula.Substring(1);
            _cell.SetCellFormula(formula);
            return this;
        }

        /// <summary>
        /// Copy style from specified cell
        /// </summary>
        /// <param name="col">Source column</param>
        /// <param name="row">Source row (1-based)</param>
        public FluentCell CopyStyleFrom(ExcelCol col, int row)
        {
            if (_cell == null) return this;

            var normalizedRow = NormalizeRow(row);
            var sourceRow = Sheet.GetRow(normalizedRow);
            var sourceCell = sourceRow?.GetCell((int)col);

            if (sourceCell?.CellStyle != null)
            {
                // Cache one clone per source style so copying from the same cell to many
                // targets reuses a single style (avoids the ~64k cell-style limit).
                string key = $"__copyfrom:{sourceCell.CellStyle.Index}";
                if (!_cellStylesCached.TryGetValue(key, out var newStyle))
                {
                    newStyle = _workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(sourceCell.CellStyle);
                    _cellStylesCached[key] = newStyle;
                }
                _cell.CellStyle = newStyle;
            }
            return this;
        }

        /// <summary>
        /// Apply an incremental style modification through the shared style cache.
        /// The cache key combines the current style's index with a deterministic operation
        /// signature, so applying the same modification to many cells that share a base style
        /// reuses one ICellStyle instead of creating one per cell. This avoids exhausting
        /// NPOI's hard ~64k cell-style-per-workbook limit on large/styled sheets.
        /// </summary>
        /// <param name="opSignature">Deterministic signature of the modification + its arguments</param>
        /// <param name="modifier">Mutates the (cloned) style; runs only on a cache miss</param>
        private FluentCell ApplyStyleModification(string opSignature, Action<ICellStyle> modifier)
        {
            if (_cell == null) return this;

            int baseIndex = _cell.CellStyle?.Index ?? -1;
            string key = $"__auto:{baseIndex}:{opSignature}";

            if (!_cellStylesCached.TryGetValue(key, out var style))
            {
                style = _workbook.CreateCellStyle();
                if (_cell.CellStyle != null)
                {
                    style.CloneStyleFrom(_cell.CellStyle);
                }
                modifier(style);
                _cellStylesCached[key] = style;
            }
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// Set background color
        /// </summary>
        /// <param name="color">Indexed color</param>
        public FluentCell SetBackgroundColor(IndexedColors color)
            => ApplyStyleModification($"bg:{color.Index}", style =>
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = color.Index;
            });

        /// <summary>
        /// Set font
        /// </summary>
        /// <param name="fontName">Font name</param>
        /// <param name="fontSize">Font size (points)</param>
        /// <param name="isBold">Is bold</param>
        public FluentCell SetFont(string? fontName = null, double? fontSize = null, bool isBold = false)
            => ApplyStyleModification($"font:{fontName}|{fontSize}|{isBold}", style =>
            {
                // CreateFont() now runs only on a cache miss, so fonts are no longer
                // created per cell (NPOI also caps fonts per workbook).
                IFont font = _workbook.CreateFont();
                if (fontName != null) font.FontName = fontName;
                if (fontSize.HasValue) font.FontHeightInPoints = fontSize.Value;
                font.IsBold = isBold;
                style.SetFont(font);
            });

        /// <summary>
        /// Set border for all sides
        /// </summary>
        /// <param name="style">Border style</param>
        public FluentCell SetBorder(BorderStyle style)
            => ApplyStyleModification($"border:{style}", s =>
            {
                s.BorderTop = style;
                s.BorderBottom = style;
                s.BorderLeft = style;
                s.BorderRight = style;
            });

        /// <summary>
        /// Set alignment
        /// </summary>
        /// <param name="horizontal">Horizontal alignment</param>
        /// <param name="vertical">Vertical alignment</param>
        public FluentCell SetAlignment(HorizontalAlignment horizontal = HorizontalAlignment.General, VerticalAlignment vertical = VerticalAlignment.Center)
            => ApplyStyleModification($"align:{horizontal}|{vertical}", style =>
            {
                style.Alignment = horizontal;
                style.VerticalAlignment = vertical;
            });

        /// <summary>
        /// Get current cell position information
        /// </summary>
        /// <returns>Column (ExcelCol) and row number (1-based)</returns>
        public (ExcelCol Column, int Row) GetPosition()
        {
            return (_col, _row + 1);  // Convert to 1-based return
        }

        /// <summary>
        /// Set number format
        /// </summary>
        /// <param name="format">Format string (e.g. "#,##0.00", "yyyy-mm-dd", "0%")</param>
        public FluentCell SetNumberFormat(string format)
        {
            if (string.IsNullOrEmpty(format)) return this;
            return ApplyStyleModification($"numfmt:{format}", style =>
            {
                IDataFormat dataFormat = _workbook.CreateDataFormat();
                style.DataFormat = dataFormat.GetFormat(format);
            });
        }

        /// <summary>
        /// Set wrap text
        /// </summary>
        /// <param name="wrap">Enable wrap text</param>
        public FluentCell SetWrapText(bool wrap = true)
            => ApplyStyleModification($"wrap:{wrap}", style => style.WrapText = wrap);

        /// <summary>
        /// Add comment
        /// </summary>
        /// <param name="text">Comment text</param>
        /// <param name="author">Author (optional)</param>
        public FluentCell SetComment(string text, string? author = null)
        {
            if (_cell == null || string.IsNullOrEmpty(text)) return this;

            ICreationHelper factory = _workbook.GetCreationHelper();
            var drawing = Sheet.CreateDrawingPatriarch();

            // Create anchor (comment display position)
            IClientAnchor anchor = factory.CreateClientAnchor();
            anchor.Col1 = _cell.ColumnIndex;
            anchor.Col2 = _cell.ColumnIndex + 2;
            anchor.Row1 = _cell.RowIndex;
            anchor.Row2 = _cell.RowIndex + 3;

            // Create comment
            IComment comment = drawing.CreateCellComment(anchor);
            comment.String = factory.CreateRichTextString(text);
            if (!string.IsNullOrEmpty(author))
            {
                comment.Author = author;
            }
            _cell.CellComment = comment;

            return this;
        }

        /// <summary>
        /// Set cell locked state (must be used with sheet protection)
        /// </summary>
        /// <param name="locked">Is locked</param>
        public FluentCell SetLocked(bool locked = true)
            => ApplyStyleModification($"locked:{locked}", style => style.IsLocked = locked);

        /// <summary>
        /// Set cell hidden formula (must be used with sheet protection)
        /// </summary>
        /// <param name="hidden">Is hidden formula</param>
        public FluentCell SetHidden(bool hidden = true)
            => ApplyStyleModification($"hidden:{hidden}", style => style.IsHidden = hidden);

        /// <summary>
        /// Set text rotation angle
        /// </summary>
        /// <param name="degrees">Rotation angle (-90 to 90)</param>
        public FluentCell SetRotation(short degrees)
            => ApplyStyleModification($"rot:{degrees}", style => style.Rotation = degrees);

        /// <summary>
        /// Set indentation level
        /// </summary>
        /// <param name="indent">Indentation level (0-15)</param>
        public FluentCell SetIndent(short indent)
            => ApplyStyleModification($"indent:{indent}", style => style.Indention = indent);
    }
}

