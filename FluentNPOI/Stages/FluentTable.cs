using NPOI.SS.UserModel;
using FluentNPOI.Base;
using FluentNPOI.Models;
using FluentNPOI.Streaming.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// Table operation class
    /// </summary>
    /// <typeparam name="T">Table data type</typeparam>
    public class FluentTable<T> : FluentSheetBase
    {
        private IEnumerable<T> _table;
        private ExcelCol _startCol;
        private int _startRow;
        private List<TableCellSet> _cellBodySets;
        private List<TableCellSet> _cellTitleSets;
        private IReadOnlyList<ColumnMapping> _columnMappings;

        public FluentTable(IWorkbook workbook, ISheet sheet, IEnumerable<T> table,
            ExcelCol startCol, int startRow,
            Dictionary<string, ICellStyle> cellStylesCached, List<TableCellSet> cellTitleSets, List<TableCellSet> cellBodySets)
            : base(workbook, sheet, cellStylesCached)
        {
            _table = table;
            // ExcelCol is already a valid column enum, no need to normalize
            _startCol = startCol;
            _startRow = NormalizeRow(startRow);
            _cellTitleSets = cellTitleSets;
            _cellBodySets = cellBodySets;
        }

        /// <summary>
        /// Use FluentMapping to set column mapping (can replace BeginTitleSet/BeginBodySet)
        /// </summary>
        public FluentTable<T> WithMapping<TMapping>(FluentMapping<TMapping> mapping) where TMapping : new()
        {
            _columnMappings = mapping.GetMappings();
            return this;
        }

        private T GetItemAt(int index)
        {
            var items = _table as IList<T> ?? _table?.ToList() ?? new List<T>();
            if (index < 0 || index >= items.Count) return default;
            return items[index];
        }

        private void SetCellAction(List<TableCellSet> cellSets, IRow rowObj, int colIndex, int targetRowIndex, object item)
        {
            foreach (var cellset in cellSets)
            {
                var cell = rowObj.GetCell(colIndex) ?? rowObj.CreateCell(colIndex);

                // Prioritize Value from TableCellNameMap, if not present, get from item
                Func<TableCellParams, object> setValueAction = cellset.SetValueAction;
                Func<TableCellParams, object> setFormulaValueAction = cellset.SetFormulaValueAction;

                TableCellParams cellParams = new TableCellParams
                {
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item
                };
                object value = cellset.CellValue ?? GetTableCellValue(cellset.CellName, item);
                cellParams.CellValue = value;

                // Prepare generic parameters (for generic delegate use)
                var cellParamsT = new TableCellParams<T>
                {
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item is T tItem ? tItem : default,
                    CellValue = value
                };

                TableCellStyleParams cellStyleParams =
                new TableCellStyleParams
                {
                    Workbook = _workbook,
                    ColNum = (ExcelCol)colIndex,
                    RowNum = targetRowIndex,
                    RowItem = item
                };
                SetCellStyle(cell, cellset, cellStyleParams);

                if (cellset.CellType == CellType.Formula)
                {
                    if (cellset.SetFormulaValueActionGeneric != null)
                    {
                        if (cellset.SetFormulaValueActionGeneric is Func<TableCellParams<T>, object> gFormula)
                        {
                            value = gFormula(cellParamsT);
                        }
                        else
                        {
                            value = cellset.SetFormulaValueActionGeneric.DynamicInvoke(cellParamsT);
                        }
                    }
                    else if (setFormulaValueAction != null)
                    {
                        value = setFormulaValueAction(cellParams);
                    }
                    SetFormulaValue(cell, value);
                }
                else
                {
                    if (cellset.SetValueActionGeneric != null)
                    {
                        if (cellset.SetValueActionGeneric is Func<TableCellParams<T>, object> gValue)
                        {
                            value = gValue(cellParamsT);
                        }
                        else
                        {
                            value = cellset.SetValueActionGeneric.DynamicInvoke(cellParamsT);
                        }
                    }
                    else if (setValueAction != null)
                    {
                        value = setValueAction(cellParams);
                    }
                    SetCellValue(cell, value, cellset.CellType);
                }

                colIndex++;
            }
        }

        private FluentTable<T> SetRow(int rowOffset = 0)
        {
            if (_cellBodySets == null || _cellBodySets.Count == 0) return this;

            var targetRowIndex = _startRow + rowOffset;

            var item = GetItemAt(rowOffset);

            int colIndex = (int)_startCol;
            if (_cellTitleSets != null && _cellTitleSets.Count > 0)
            {
                var titleRowObj = _sheet.GetRow(_startRow) ?? _sheet.CreateRow(_startRow);
                SetCellAction(_cellTitleSets, titleRowObj, colIndex, _startRow, item);
                targetRowIndex++;
            }

            var rowObj = _sheet.GetRow(targetRowIndex) ?? _sheet.CreateRow(targetRowIndex);
            SetCellAction(_cellBodySets, rowObj, colIndex, targetRowIndex, item);

            return this;
        }

        /// <summary>
        /// Write a row using FluentMapping
        /// </summary>
        private FluentTable<T> SetRowWithMapping(int rowOffset, bool writeTitle)
        {
            if (_columnMappings == null || _columnMappings.Count == 0) return this;

            var item = GetItemAt(rowOffset);
            var targetRowIndex = _startRow + rowOffset + (writeTitle ? 1 : 0);

            // Write title row (only for the first time)
            if (writeTitle && rowOffset == 0)
            {
                foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
                {
                    // Calculate title row position for this column (add per-column offset)
                    var titleRowIndex = _startRow + map.RowOffset;
                    var titleRow = _sheet.GetRow(titleRowIndex) ?? _sheet.CreateRow(titleRowIndex);
                    var colIdx = (int)map.ColumnIndex.Value;
                    var cell = titleRow.GetCell(colIdx) ?? titleRow.CreateCell(colIdx);
                    cell.SetCellValue(map.Title ?? map.Property?.Name ?? map.ColumnName ?? "");

                    // Apply title style
                    if (map.TitleStyleRef != null)
                    {
                        string refKey = $"{_sheet.SheetName}_{map.TitleStyleRef.Column}{map.TitleStyleRef.Row}";
                        if (_cellStylesCached.TryGetValue(refKey, out var cachedStyle))
                        {
                            cell.CellStyle = cachedStyle;
                        }
                        else
                        {
                            var refRow = _sheet.GetRow(map.TitleStyleRef.Row);
                            var refCell = refRow?.GetCell((int)map.TitleStyleRef.Column);
                            if (refCell != null && refCell.CellStyle != null)
                            {
                                ICellStyle newCellStyle = _workbook.CreateCellStyle();
                                newCellStyle.CloneStyleFrom(refCell.CellStyle);
                                _cellStylesCached[refKey] = newCellStyle;
                                cell.CellStyle = newCellStyle;
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(map.TitleStyleKey) && _cellStylesCached.TryGetValue(map.TitleStyleKey, out var titleStyle))
                    {
                        cell.CellStyle = titleStyle;
                    }
                    else if (cell.CellStyle.Index == 0) // Try global style
                    {
                        // Prioritize Sheet global style, if not present, use Workbook global style
                        string sheetGlobalKey = $"global_{_sheet.SheetName}";
                        if (_cellStylesCached.TryGetValue(sheetGlobalKey, out var sheetGlobalStyle))
                        {
                            cell.CellStyle = sheetGlobalStyle;
                        }
                        else if (_cellStylesCached.TryGetValue("global", out var workbookGlobalStyle))
                        {
                            cell.CellStyle = workbookGlobalStyle;
                        }
                    }
                }
            }

            // Write data row
            foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
            {
                // Calculate data row position for this column (add per-column offset)
                var colTargetRowIndex = targetRowIndex + map.RowOffset;
                var dataRow = _sheet.GetRow(colTargetRowIndex) ?? _sheet.CreateRow(colTargetRowIndex);
                var colIdx = (int)map.ColumnIndex.Value;
                var cell = dataRow.GetCell(colIdx) ?? dataRow.CreateCell(colIdx);

                // Formula first
                if (map.FormulaFunc != null)
                {
                    var formula = map.FormulaFunc(targetRowIndex + 1, (ExcelCol)colIdx); // Excel 是 1-based
                    cell.SetCellFormula(formula);
                }
                else
                {
                    // Calculate value (Prioritize ValueFunc, otherwise get from property)
                    object value;
                    if (map.ValueFunc != null)
                    {
                        value = map.ValueFunc(item, targetRowIndex + 1, (ExcelCol)colIdx);
                    }
                    else if (map.Property != null)
                    {
                        value = map.Property.GetValue(item);
                    }
                    else
                    {
                        value = null;
                    }

                    SetCellValue(cell, value, map.CellType ?? CellType.Unknown);
                }

                // Register automatically generated style
                if (!string.IsNullOrEmpty(map.GeneratedStyleKey) && map.StyleConfig != null)
                {
                    RegisterStyle(map.GeneratedStyleKey, map.StyleConfig);
                }

                // Apply data style
                string styleKey = map.StyleKey;

                // If StyleKey is not specified, try to use GeneratedStyleKey
                if (string.IsNullOrEmpty(styleKey))
                {
                    styleKey = map.GeneratedStyleKey;
                }

                // If there is dynamic style setting, use it first
                if (map.DynamicStyleFunc != null)
                {
                    string dynamicKey = map.DynamicStyleFunc(item);
                    if (!string.IsNullOrEmpty(dynamicKey))
                    {
                        styleKey = dynamicKey;
                    }
                }

                if (!string.IsNullOrEmpty(styleKey) && _cellStylesCached.TryGetValue(styleKey, out var dataStyle))
                {
                    cell.CellStyle = dataStyle;
                }

                // Copy data style
                if (map.DataStyleRef != null)
                {
                    string refKey = $"{_sheet.SheetName}_{map.DataStyleRef.Column}{map.DataStyleRef.Row}";
                    if (_cellStylesCached.TryGetValue(refKey, out var cachedStyle))
                    {
                        cell.CellStyle = cachedStyle;
                    }
                    else
                    {
                        var refRow = _sheet.GetRow(map.DataStyleRef.Row);
                        var refCell = refRow?.GetCell((int)map.DataStyleRef.Column);
                        if (refCell != null && refCell.CellStyle != null)
                        {
                            ICellStyle newCellStyle = _workbook.CreateCellStyle();
                            newCellStyle.CloneStyleFrom(refCell.CellStyle);
                            _cellStylesCached[refKey] = newCellStyle;
                            cell.CellStyle = newCellStyle;
                        }
                    }
                }

                // If no style is set, try to apply global style
                // Prioritize Sheet global style, if not present, use Workbook global style
                if (cell.CellStyle.Index == 0) // Default style index is usually 0
                {
                    string sheetGlobalKey = $"global_{_sheet.SheetName}";
                    if (_cellStylesCached.TryGetValue(sheetGlobalKey, out var sheetGlobalStyle))
                    {
                        cell.CellStyle = sheetGlobalStyle;
                    }
                    else if (_cellStylesCached.TryGetValue("global", out var workbookGlobalStyle))
                    {
                        cell.CellStyle = workbookGlobalStyle;
                    }
                }
            }

            return this;
        }


        public FluentTable<T> BuildRows()
        {
            // If FluentMapping exists, write using mapping
            if (_columnMappings != null)
            {
                bool writeTitle = _columnMappings.Any(m => !string.IsNullOrEmpty(m.Title));
                for (int i = 0; i < _table.Count(); i++)
                {
                    SetRowWithMapping(i, writeTitle);
                }
                return this;
            }

            // Otherwise use original method
            for (int i = 0; i < _table.Count(); i++)
            {
                SetRow(i);
            }
            return this;
        }

        /// <summary>
        /// Get data row count (excluding title)
        /// </summary>
        public int RowCount => _table.Count();

        /// <summary>
        /// Get column count
        /// </summary>
        public int ColumnCount => _columnMappings?.Count ?? 0;

        /// <summary>
        /// Auto size all columns
        /// </summary>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> AutoSizeColumns()
        {
            if (_columnMappings == null) return this;

            foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
            {
                _sheet.AutoSizeColumn((int)map.ColumnIndex.Value);
            }
            return this;
        }

        /// <summary>
        /// Set same width for all columns
        /// </summary>
        /// <param name="width">Width (characters)</param>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> SetColumnWidths(int width)
        {
            if (_columnMappings == null) return this;

            foreach (var map in _columnMappings.Where(m => m.ColumnIndex.HasValue))
            {
                _sheet.SetColumnWidth((int)map.ColumnIndex.Value, width * 256);
            }
            return this;
        }

        /// <summary>
        /// Set title row height
        /// </summary>
        /// <param name="heightInPoints">Row height (points)</param>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> SetTitleRowHeight(float heightInPoints)
        {
            var row = _sheet.GetRow(_startRow);
            if (row != null)
            {
                row.HeightInPoints = heightInPoints;
            }
            return this;
        }

        /// <summary>
        /// Set data row height
        /// </summary>
        /// <param name="heightInPoints">Row height (points)</param>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> SetDataRowHeights(float heightInPoints)
        {
            int dataStartRow = _startRow + 1; // Skip title row
            for (int i = 0; i < _table.Count(); i++)
            {
                var row = _sheet.GetRow(dataStartRow + i);
                if (row != null)
                {
                    row.HeightInPoints = heightInPoints;
                }
            }
            return this;
        }

        /// <summary>
        /// Freeze title row (keep title row fixed when scrolling)
        /// </summary>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> FreezeTitleRow()
        {
            // Freeze below title row (_startRow + 1)
            _sheet.CreateFreezePane(0, _startRow + 1);
            return this;
        }

        /// <summary>
        /// Set auto filter
        /// </summary>
        /// <returns>FluentTable instance, supports method chaining</returns>
        public FluentTable<T> SetAutoFilter()
        {
            if (_columnMappings == null || !_columnMappings.Any()) return this;

            var columns = _columnMappings.Where(m => m.ColumnIndex.HasValue).ToList();
            if (!columns.Any()) return this;

            int firstCol = (int)columns.Min(m => m.ColumnIndex.Value);
            int lastCol = (int)columns.Max(m => m.ColumnIndex.Value);
            int lastRow = _startRow + _table.Count(); // Title row + data row count

            var range = new NPOI.SS.Util.CellRangeAddress(_startRow, lastRow, firstCol, lastCol);
            _sheet.SetAutoFilter(range);
            return this;
        }

        /// <summary>
        /// Get table range (for setting styles, etc.)
        /// </summary>
        /// <returns>Start row, end row, start column, end column (all 0-based)</returns>
        public (int StartRow, int EndRow, int StartCol, int EndCol) GetTableRange()
        {
            if (_columnMappings == null || !_columnMappings.Any())
                return (_startRow, _startRow, 0, 0);

            var columns = _columnMappings.Where(m => m.ColumnIndex.HasValue).ToList();
            if (!columns.Any())
                return (_startRow, _startRow, 0, 0);

            int firstCol = (int)columns.Min(m => m.ColumnIndex.Value);
            int lastCol = (int)columns.Max(m => m.ColumnIndex.Value);
            int lastRow = _startRow + _table.Count();

            return (_startRow, lastRow, firstCol, lastCol);
        }
    }
}
