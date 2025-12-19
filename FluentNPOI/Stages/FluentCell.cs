using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Helpers;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 單元格操作類
    /// </summary>
    public class FluentCell : FluentCellBase
    {
        private ICell _cell;
        private ExcelCol _col;
        private int _row;
        public FluentCell(IWorkbook workbook, ISheet sheet,
        ICell cell, Dictionary<string, ICellStyle> cellStylesCached = null)
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

        public FluentCell SetCellStyle(Func<TableCellStyleParams, CellStyleConfig> cellStyleAction)
        {
            if (_cell == null) return this;

            var cellStyleParams = new TableCellStyleParams
            {
                Workbook = _workbook,
                ColNum = (ExcelCol)_cell.ColumnIndex,
                RowNum = _cell.RowIndex,
                RowItem = null
            };

            // ✅ 先調用函數獲取樣式配置
            var config = cellStyleAction(cellStyleParams);

            if (!string.IsNullOrWhiteSpace(config.Key))
            {
                // ✅ 先檢查緩存
                if (!_cellStylesCached.ContainsKey(config.Key))
                {
                    // ✅ 只在不存在時才創建新樣式
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    config.StyleSetter(newCellStyle);
                    _cellStylesCached.Add(config.Key, newCellStyle);
                }
                _cell.CellStyle = _cellStylesCached[config.Key];
            }
            else
            {
                // 如果沒有返回 key，創建臨時樣式（不緩存）
                ICellStyle tempStyle = _workbook.CreateCellStyle();
                config.StyleSetter(tempStyle);
                _cell.CellStyle = tempStyle;
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
        /// 獲取當前單元格的值
        /// </summary>
        /// <returns>單元格的值（根據類型返回 bool, DateTime, double, string 或 null）</returns>
        public object GetValue()
        {
            return GetCellValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的值並轉換為指定類型
        /// </summary>
        /// <typeparam name="T">目標類型</typeparam>
        /// <returns>轉換後的值</returns>
        public T GetValue<T>()
        {
            return GetCellValue<T>(_cell);
        }

        /// <summary>
        /// 獲取當前單元格的公式字符串（如果是公式單元格）
        /// </summary>
        /// <returns>公式字符串（不含 '=' 前綴），如果不是公式則返回 null</returns>
        public string GetFormula()
        {
            return GetCellFormulaValue(_cell);
        }

        /// <summary>
        /// 獲取當前單元格對象
        /// </summary>
        /// <returns>NPOI ICell 對象</returns>
        public ICell GetCell()
        {
            return _cell;
        }

        /// <summary>
        /// 在單元格中設置圖片（自動計算高度，保持原圖比例）
        /// </summary>
        /// <param name="pictureBytes">圖片字節數組</param>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <param name="columnWidthRatio">列寬轉換比例（默認 7.0，表示像素寬度除以該值得到 Excel 列寬字符數）</param>
        /// <returns>FluentCell 實例，支持鏈式調用</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, AnchorType anchorType = AnchorType.MoveAndResize, double columnWidthRatio = 7.0)
        {
            // 自動計算高度（需要從圖片中讀取原始尺寸）
            // 由於無法直接從字節數組獲取圖片尺寸，這裡使用寬度作為高度（1:1比例）
            // 如果需要更精確的比例，可以考慮使用 System.Drawing.Image 或其他圖像庫
            return SetPictureOnCell(pictureBytes, imgWidth, imgWidth, anchorType, columnWidthRatio);
        }

        /// <summary>
        /// 在單元格中設置圖片（手動設置寬度和高度）
        /// </summary>
        /// <param name="pictureBytes">圖片字節數組</param>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="imgHeight">圖片高度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <param name="columnWidthRatio">列寬轉換比例（默認 7.0，表示像素寬度除以該值得到 Excel 列寬字符數）</param>
        /// <param name="pictureAction">圖片操作委托</param>
        /// <returns>FluentCell 實例，支持鏈式調用</returns>
        public FluentCell SetPictureOnCell(byte[] pictureBytes, int imgWidth, int imgHeight, AnchorType anchorType = AnchorType.MoveAndResize,
        double columnWidthRatio = 7.0, Action<IPicture> pictureAction = null)
        {
            // 參數驗證
            ValidatePictureParameters(pictureBytes, imgWidth, imgHeight, columnWidthRatio);

            // 設置列寬
            double columnWidth = CalculateColumnWidth(imgWidth, columnWidthRatio);
            _sheet.SetColumnWidth((int)_col, (int)Math.Round(columnWidth * 256));

            // 獲取圖片類型並添加到工作簿
            var picType = GetPictureType(pictureBytes);
            int picIndex = _workbook.AddPicture(pictureBytes, picType);

            // 創建繪圖對象和錨點
            IDrawing drawing = _sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = CreatePictureAnchor(imgWidth, imgHeight, anchorType);

            // 創建圖片
            IPicture pict = drawing.CreatePicture(anchor, picIndex);

            pictureAction?.Invoke(pict);

            return this;
        }

        /// <summary>
        /// 驗證圖片參數
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
        /// 計算列寬（將像素寬度轉換為 Excel 列寬單位）
        /// </summary>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="columnWidthRatio">轉換比例</param>
        /// <returns>Excel 列寬（字符數）</returns>
        private double CalculateColumnWidth(int imgWidth, double columnWidthRatio)
        {
            // Excel 列寬單位：1 個字符寬度 = 256 單位
            // 將像素寬度除以轉換比例得到字符數
            return imgWidth / columnWidthRatio;
        }

        /// <summary>
        /// 創建圖片錨點，設置完整的位置和大小信息
        /// </summary>
        /// <param name="imgWidth">圖片寬度（像素）</param>
        /// <param name="imgHeight">圖片高度（像素）</param>
        /// <param name="anchorType">錨點類型</param>
        /// <returns>配置好的 IClientAnchor 對象</returns>
        private IClientAnchor CreatePictureAnchor(int imgWidth, int imgHeight, AnchorType anchorType)
        {
            ICreationHelper creationHelper = _workbook.GetCreationHelper();
            IClientAnchor anchor = creationHelper.CreateClientAnchor();

            // 設置起始位置（_row 已經是 0-based，因為在 SetCellPosition 中已經轉換）
            anchor.Col1 = (short)_col;
            anchor.Row1 = (short)_row;

            // 計算結束位置（Col2 和 Row2）
            // 根據圖片尺寸和單元格大小計算需要跨越多少列和行
            // Excel 默認列寬約為 8.43 字符（約 64 像素），行高約為 15 像素
            // 這裡使用簡化的計算方式

            // 獲取當前列寬（以字符為單位）
            // GetColumnWidth 返回 int（以 1/256 字符為單位），轉換為字符數
            double columnWidthInChars = _sheet.GetColumnWidth((int)_col) / 256.0;

            // 獲取當前行高（以點為單位，1 點 ≈ 1.33 像素）
            IRow row = _sheet.GetRow(_row) ?? _sheet.CreateRow(_row);
            short rowHeightInPoints = row.Height > 0 ? (short)(row.Height / 20.0) : (short)15; // 默認行高約 15 點

            // 計算需要跨越的列數（考慮列寬）
            // 假設 1 字符寬度 ≈ 7 像素（可根據實際情況調整）
            double pixelsPerChar = 7.0;
            double columnsNeeded = imgWidth / (columnWidthInChars * pixelsPerChar);
            short col2 = (short)Math.Min((int)_col + (int)Math.Ceiling(columnsNeeded), 16383); // Excel 最大列數限制

            // 計算需要跨越的行數（考慮行高）
            // 1 點 ≈ 1.33 像素
            double pixelsPerPoint = 1.33;
            double rowsNeeded = imgHeight / (rowHeightInPoints * pixelsPerPoint);
            short row2 = (short)Math.Min(_row + (int)Math.Ceiling(rowsNeeded), 1048575); // Excel 最大行數限制

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

            // EMF: 01 00 00 00 (但需要更多检查，EMF 文件通常以这个开头)
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x00 && pictureBytes[3] == 0x00)
            {
                // 检查是否是有效的 EMF 文件（EMF 文件头通常是 40 字节）
                if (pictureBytes.Length >= 40)
                {
                    // EMF 文件的第二个 DWORD 应该是文件大小
                    // 这里做简单检查，如果符合 EMF 特征就返回 EMF
                    return PictureType.EMF;
                }
            }

            // WMF: 通常以 01 00 09 00 开头（但需要更多检查）
            if (pictureBytes.Length >= 4 &&
                pictureBytes[0] == 0x01 && pictureBytes[1] == 0x00 && pictureBytes[2] == 0x09 && pictureBytes[3] == 0x00)
            {
                return PictureType.WMF;
            }

            throw new NotSupportedException($"Unsupported picture format. File header: {BitConverter.ToString(pictureBytes, 0, Math.Min(8, pictureBytes.Length))}");
        }

        /// <summary>
        /// 設定當前操作的儲存格位置
        /// </summary>
        /// <param name="col">欄位置</param>
        /// <param name="row">列位置（1-based）</param>
        public FluentCell SetCellPosition(ExcelCol col, int row)
        {
            _cell = SetCellPositionInternal(col, row);
            _col = col;
            _row = NormalizeRow(row);  // 存儲 0-based 的 row
            return this;
        }

        /// <summary>
        /// 設定公式（不含 '=' 前綴）
        /// </summary>
        /// <param name="formula">公式字串（例如 "SUM(A1:A10)"）</param>
        public FluentCell SetFormula(string formula)
        {
            if (_cell == null) return this;
            if (string.IsNullOrWhiteSpace(formula)) return this;

            // 移除 '=' 前綴（如果有的話）
            if (formula.StartsWith("=")) formula = formula.Substring(1);
            _cell.SetCellFormula(formula);
            return this;
        }

        /// <summary>
        /// 從指定儲存格複製樣式
        /// </summary>
        /// <param name="col">來源欄</param>
        /// <param name="row">來源列（1-based）</param>
        public FluentCell CopyStyleFrom(ExcelCol col, int row)
        {
            if (_cell == null) return this;

            var normalizedRow = NormalizeRow(row);
            var sourceRow = _sheet.GetRow(normalizedRow);
            var sourceCell = sourceRow?.GetCell((int)col);

            if (sourceCell?.CellStyle != null)
            {
                ICellStyle newStyle = _workbook.CreateCellStyle();
                newStyle.CloneStyleFrom(sourceCell.CellStyle);
                _cell.CellStyle = newStyle;
            }
            return this;
        }

        /// <summary>
        /// 設定背景顏色
        /// </summary>
        /// <param name="color">索引顏色</param>
        public FluentCell SetBackgroundColor(IndexedColors color)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = color.Index;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定字體
        /// </summary>
        /// <param name="fontName">字體名稱</param>
        /// <param name="fontSize">字體大小（點）</param>
        /// <param name="isBold">是否粗體</param>
        public FluentCell SetFont(string fontName = null, double? fontSize = null, bool isBold = false)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }

            IFont font = _workbook.CreateFont();
            if (fontName != null) font.FontName = fontName;
            if (fontSize.HasValue) font.FontHeightInPoints = fontSize.Value;
            font.IsBold = isBold;
            style.SetFont(font);
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定四邊邊框
        /// </summary>
        /// <param name="style">邊框樣式</param>
        public FluentCell SetBorder(BorderStyle style)
        {
            if (_cell == null) return this;

            ICellStyle cellStyle = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                cellStyle.CloneStyleFrom(_cell.CellStyle);
            }
            cellStyle.BorderTop = style;
            cellStyle.BorderBottom = style;
            cellStyle.BorderLeft = style;
            cellStyle.BorderRight = style;
            _cell.CellStyle = cellStyle;
            return this;
        }

        /// <summary>
        /// 設定對齊方式
        /// </summary>
        /// <param name="horizontal">水平對齊</param>
        /// <param name="vertical">垂直對齊</param>
        public FluentCell SetAlignment(HorizontalAlignment horizontal = HorizontalAlignment.General, VerticalAlignment vertical = VerticalAlignment.Center)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.Alignment = horizontal;
            style.VerticalAlignment = vertical;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 取得當前儲存格的位置資訊
        /// </summary>
        /// <returns>欄位（ExcelCol）和列號（1-based）</returns>
        public (ExcelCol Column, int Row) GetPosition()
        {
            return (_col, _row + 1);  // 轉換為 1-based 返回
        }

        /// <summary>
        /// 設定數值格式
        /// </summary>
        /// <param name="format">格式字串（例如 "#,##0.00", "yyyy-mm-dd", "0%"）</param>
        public FluentCell SetNumberFormat(string format)
        {
            if (_cell == null || string.IsNullOrEmpty(format)) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }

            IDataFormat dataFormat = _workbook.CreateDataFormat();
            style.DataFormat = dataFormat.GetFormat(format);
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定自動換行
        /// </summary>
        /// <param name="wrap">是否啟用自動換行</param>
        public FluentCell SetWrapText(bool wrap = true)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.WrapText = wrap;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 添加備註（批註）
        /// </summary>
        /// <param name="text">備註文字</param>
        /// <param name="author">作者（可選）</param>
        public FluentCell SetComment(string text, string author = null)
        {
            if (_cell == null || string.IsNullOrEmpty(text)) return this;

            ICreationHelper factory = _workbook.GetCreationHelper();
            IDrawing drawing = _sheet.CreateDrawingPatriarch();

            // 創建錨點（備註顯示位置）
            IClientAnchor anchor = factory.CreateClientAnchor();
            anchor.Col1 = _cell.ColumnIndex;
            anchor.Col2 = _cell.ColumnIndex + 2;
            anchor.Row1 = _cell.RowIndex;
            anchor.Row2 = _cell.RowIndex + 3;

            // 創建備註
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
        /// 設定儲存格鎖定狀態（需配合保護工作表使用）
        /// </summary>
        /// <param name="locked">是否鎖定</param>
        public FluentCell SetLocked(bool locked = true)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.IsLocked = locked;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定儲存格隱藏公式（需配合保護工作表使用）
        /// </summary>
        /// <param name="hidden">是否隱藏公式</param>
        public FluentCell SetHidden(bool hidden = true)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.IsHidden = hidden;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定文字旋轉角度
        /// </summary>
        /// <param name="degrees">旋轉角度（-90 到 90）</param>
        public FluentCell SetRotation(short degrees)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.Rotation = degrees;
            _cell.CellStyle = style;
            return this;
        }

        /// <summary>
        /// 設定縮進層級
        /// </summary>
        /// <param name="indent">縮進層級（0-15）</param>
        public FluentCell SetIndent(short indent)
        {
            if (_cell == null) return this;

            ICellStyle style = _workbook.CreateCellStyle();
            if (_cell.CellStyle != null)
            {
                style.CloneStyleFrom(_cell.CellStyle);
            }
            style.Indention = indent;
            _cell.CellStyle = style;
            return this;
        }
    }
}

