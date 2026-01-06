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
    /// Sheet operation class
    /// </summary>
    public class FluentSheet : FluentCellBase
    {
        /// <summary>
        /// Initialize FluentSheet instance
        /// </summary>
        /// <param name="workbook">Workbook object</param>
        /// <param name="sheet">Sheet object</param>
        /// <param name="cellStylesCached">Style cache dictionary</param>
        public FluentSheet(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
            : base(workbook, sheet, cellStylesCached)
        {
        }

        /// <summary>
        /// Get NPOI Sheet object
        /// </summary>
        /// <returns>ISheet object</returns>
        public ISheet GetSheet()
        {
            return _sheet;
        }

        /// <summary>
        /// Set column width
        /// </summary>
        /// <param name="col">Column position</param>
        /// <param name="width">Width (characters)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetColumnWidth(ExcelCol col, int width)
        {
            _sheet.SetColumnWidth((int)col, width * 256);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Batch set column width
        /// </summary>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="width">Width (characters)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetColumnWidth(ExcelCol startCol, ExcelCol endCol, int width)
        {
            for (int i = (int)startCol; i <= (int)endCol; i++)
            {
                _sheet.SetColumnWidth(i, width * 256);
            }
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Set row height
        /// </summary>
        /// <param name="row">Row index (1-based)</param>
        /// <param name="heightInPoints">Height (points)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetRowHeight(int row, float heightInPoints)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow) ?? _sheet.CreateRow(normalizedRow);
            rowObj.HeightInPoints = heightInPoints;
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Batch set row height
        /// </summary>
        /// <param name="startRow">Start row (1-based)</param>
        /// <param name="endRow">End row (1-based)</param>
        /// <param name="heightInPoints">Height (points)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetRowHeight(int startRow, int endRow, float heightInPoints)
        {
            var normalizedStartRow = NormalizeRow(startRow);
            var normalizedEndRow = NormalizeRow(endRow);
            for (int i = normalizedStartRow; i <= normalizedEndRow; i++)
            {
                var rowObj = _sheet.GetRow(i) ?? _sheet.CreateRow(i);
                rowObj.HeightInPoints = heightInPoints;
            }
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Set default row height
        /// </summary>
        /// <param name="heightInPoints">Height (points)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetDefaultRowHeight(float heightInPoints)
        {
            _sheet.DefaultRowHeightInPoints = heightInPoints;
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Set default column width
        /// </summary>
        /// <param name="width">Width (characters)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetDefaultColumnWidth(int width)
        {
            _sheet.DefaultColumnWidth = width;
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Merge cells (Horizontal)
        /// </summary>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="row">Row index (1-based)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetExcelCellMerge(ExcelCol startCol, ExcelCol endCol, int row)
        {
            _sheet.SetExcelCellMerge(startCol, endCol, row);
            return new FluentSheet(_workbook, _sheet, _cellStylesCached);
        }

        /// <summary>
        /// Merge cells (Region)
        /// </summary>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="firstRow">First row (1-based)</param>
        /// <param name="lastRow">Last row (1-based)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetExcelCellMerge(ExcelCol startCol, ExcelCol endCol, int firstRow, int lastRow)
        {
            var region = new NPOI.SS.Util.CellRangeAddress(firstRow - 1, lastRow - 1, (int)startCol, (int)endCol);
            _sheet.AddMergedRegion(region);
            return this;
        }

        /// <summary>
        /// Freeze Pane (Split window)
        /// </summary>
        /// <param name="colSplit">Horizontal split position (column count)</param>
        /// <param name="rowSplit">Vertical split position (row count)</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        /// <returns>FluentSheet instance</returns>
        public FluentSheet CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn = 0, int topRow = 0)
        {
            _sheet.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
            return this;
        }

        /// <summary>
        /// Freeze Title Row (Fix first row)
        /// </summary>
        /// <param name="rowCount">Number of rows to freeze (default 1)</param>
        /// <returns>FluentSheet instance</returns>
        public FluentSheet FreezeTitleRow(int rowCount = 1)
        {
            _sheet.CreateFreezePane(0, rowCount);
            return this;
        }

        /// <summary>
        /// Set cell position and return FluentCell instance
        /// </summary>
        /// <param name="col">Column position</param>
        /// <param name="row">Row index (1-based)</param>
        /// <returns>FluentCell instance</returns>
        public FluentCell SetCellPosition(ExcelCol col, int row)
        {
            var cell = SetCellPositionInternal(col, row);
            return new FluentCell(_workbook, _sheet, cell, _cellStylesCached);
        }

        /// <summary>
        /// Use FluentMapping to set table data (Recommended)
        /// </summary>
        /// <typeparam name="T">Data Type</typeparam>
        /// <param name="table">Data collection</param>
        /// <param name="mapping">FluentMapping configuration</param>
        /// <param name="startRow">Start row (1-based), defaults to mapping.StartRow (default 1) if not specified</param>
        /// <returns>FluentTable instance (mapping applied)</returns>
        public FluentTable<T> SetTable<T>(IEnumerable<T> table, FluentMapping<T> mapping, int? startRow = null) where T : new()
        {
            // Use startRow from parameter first, otherwise use StartRow from mapping
            int actualStartRow = startRow ?? mapping.StartRow;
            var fluentTable = new FluentTable<T>(_workbook, _sheet, table, ExcelCol.A, actualStartRow, _cellStylesCached);
            return fluentTable.WithMapping(mapping);
        }

        /// <summary>
        /// Use DataTableMapping to write DataTable data
        /// </summary>
        /// <param name="dataTable">DataTable data</param>
        /// <param name="mapping">DataTableMapping configuration (can be null, will be automatically generated)</param>
        /// <param name="startRow">Start row (1-based), defaults to mapping.StartRow (default 1) if not specified</param>
        public FluentSheet WriteDataTable(System.Data.DataTable dataTable, DataTableMapping mapping = null, int? startRow = null)
        {
            var actualMapping = mapping ?? DataTableMapping.FromDataTable(dataTable);
            // Use startRow from parameter first, otherwise use StartRow from mapping
            int actualStartRow = startRow ?? actualMapping.StartRow;
            var mappings = actualMapping.GetMappings().Where(m => m.ColumnIndex.HasValue).ToList();

            // Register auto-generated styles
            foreach (var map in mappings)
            {
                if (map.StyleConfig != null && !string.IsNullOrEmpty(map.GeneratedStyleKey))
                {
                    RegisterStyle(map.GeneratedStyleKey, map.StyleConfig);
                }
                if (map.TitleStyleConfig != null && !string.IsNullOrEmpty(map.GeneratedTitleStyleKey))
                {
                    RegisterStyle(map.GeneratedTitleStyleKey, map.TitleStyleConfig);
                }
            }

            bool writeTitle = mappings.Any(m => !string.IsNullOrEmpty(m.Title));

            // Write title row
            if (writeTitle)
            {
                var titleRow = GetOrCreateRow(actualStartRow - 1);
                foreach (var map in mappings)
                {
                    var cell = GetOrCreateCell(titleRow, (int)map.ColumnIndex.Value);
                    cell.SetCellValue(map.Title ?? map.ColumnName ?? "");
                    ApplyStyle(cell, map.TitleStyleKey ?? map.GeneratedTitleStyleKey);
                }
            }

            // Write data rows
            int dataRowStart = actualStartRow - 1 + (writeTitle ? 1 : 0);
            for (int rowIdx = 0; rowIdx < dataTable.Rows.Count; rowIdx++)
            {
                var dataRow = dataTable.Rows[rowIdx];
                var excelRow = GetOrCreateRow(dataRowStart + rowIdx);

                foreach (var map in mappings)
                {
                    var colIdx = (int)map.ColumnIndex.Value;
                    var cell = GetOrCreateCell(excelRow, colIdx);

                    if (map.FormulaFunc != null)
                        cell.SetCellFormula(map.FormulaFunc(dataRowStart + rowIdx + 1, (ExcelCol)colIdx));
                    else
                        SetCellValueInternal(cell, actualMapping.GetValue(map, dataRow, dataRowStart + rowIdx + 1, (ExcelCol)colIdx));

                    // Apply dynamic style from function if present, otherwise use static/generated style
                    string styleKey = null;
                    if (map.DynamicStyleFunc != null)
                    {
                        styleKey = map.DynamicStyleFunc(dataRow);
                    }
                    
                    ApplyStyle(cell, styleKey ?? map.StyleKey ?? map.GeneratedStyleKey);
                }
            }

            return this;
        }

        private IRow GetOrCreateRow(int rowIndex) => _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);

        private ICell GetOrCreateCell(IRow row, int colIndex) => row.GetCell(colIndex) ?? row.CreateCell(colIndex);

        private void ApplyStyle(ICell cell, string styleKey)
        {
            if (string.IsNullOrEmpty(styleKey)) styleKey = "global";

            if (!string.IsNullOrEmpty(styleKey) && _cellStylesCached.TryGetValue(styleKey, out var style))
                cell.CellStyle = style;
        }

        private void SetCellValueInternal(ICell cell, object value)
        {
            if (value == null || value == DBNull.Value) { cell.SetCellValue(""); return; }

            switch (value)
            {
                case string s: cell.SetCellValue(s); break;
                case DateTime dt: cell.SetCellValue(dt); break;
                case double d: cell.SetCellValue(d); break;
                case int i: cell.SetCellValue(i); break;
                case long l: cell.SetCellValue(l); break;
                case decimal dec: cell.SetCellValue((double)dec); break;
                case bool b: cell.SetCellValue(b); break;
                default: cell.SetCellValue(value.ToString()); break;
            }
        }

        /// <summary>
        /// Get cell value and convert to specified type
        /// </summary>
        private object GetCellValueForType(ICell cell, System.Type targetType)
        {
            if (cell == null)
                return null;

            try
            {
                // Use generic method to get value - find protected GetCellValue<T>(ICell) method
                var methods = typeof(FluentCellBase).GetMethods(
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                var method = methods.FirstOrDefault(m =>
                    m.Name == "GetCellValue" &&
                    m.IsGenericMethodDefinition &&
                    m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType == typeof(ICell));

                if (method != null)
                {
                    var genericMethod = method.MakeGenericMethod(targetType);
                    return genericMethod.Invoke(this, new object[] { cell });
                }
            }
            catch
            {
                // If reflection fails, use fallback
            }

            // Fallback: get value directly
            return GetCellValue(cell);
        }

        /// <summary>
        /// Get cell value at specified position
        /// </summary>
        /// <param name="col">Column position</param>
        /// <param name="row">Row position (1-based)</param>
        /// <returns>Cell value</returns>
        public object GetCellValue(ExcelCol col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            return GetCellValue(cell);
        }

        /// <summary>
        /// Get cell value at specified position and convert to specified type
        /// </summary>
        /// <typeparam name="T">Target type</typeparam>
        /// <param name="col">Column position</param>
        /// <param name="row">Row position (1-based)</param>
        /// <returns>Converted value</returns>
        public T GetCellValue<T>(ExcelCol col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return default;

            var cell = rowObj.GetCell((int)col);
            return GetCellValue<T>(cell);
        }

        /// <summary>
        /// Get cell formula string at specified position
        /// </summary>
        /// <param name="col">Column position</param>
        /// <param name="row">Row position (1-based)</param>
        /// <returns>Formula string (without '=' prefix)</returns>
        public string GetCellFormula(ExcelCol col, int row)
        {
            var normalizedRow = NormalizeRow(row);
            var rowObj = _sheet.GetRow(normalizedRow);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell((int)col);
            return GetCellFormulaValue(cell);
        }

        // /// <summary>
        // /// 獲取指定位置的單元格對象（用於更高級的讀取操作）
        // /// </summary>
        // /// <param name="col">列位置</param>
        // /// <param name="row">行位置（1-based）</param>
        // /// <returns>FluentCell 對象，可以鏈式調用讀取方法</returns>
        // public FluentCell GetCellPosition(ExcelCol col, int row)
        // {
        //     var normalizedRow = NormalizeRow(row);
        //     var rowObj = _sheet.GetRow(normalizedRow);
        //     if (rowObj == null) return null;

        //     var cell = rowObj.GetCell((int)col);
        //     if (cell == null) return null;

        //     return new FluentCell(_workbook, _sheet, cell, col, normalizedRow, _cellStylesCached);
        // }

        /// <summary>
        /// Batch set cell range style (using style key)
        /// </summary>
        /// <param name="cellStyleKey">Style cache key</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="startRow">Start row (1-based)</param>
        /// <param name="endRow">End row (1-based)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public FluentSheet SetCellStyleRange(string cellStyleKey, ExcelCol startCol, ExcelCol endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(new CellStyleConfig(cellStyleKey, null), startCol, endCol, startRow, endRow);
            return this;
        }

        /// <summary>
        /// Batch set cell range style (using style configuration)
        /// </summary>
        /// <param name="cellStyleConfig">Style configuration object</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="startRow">Start row (1-based)</param>
        /// <param name="endRow">End row (1-based)</param>
        /// <returns>FluentSheet instance, supports method chaining</returns>
        public new FluentSheet SetCellStyleRange(CellStyleConfig cellStyleConfig, ExcelCol startCol, ExcelCol endCol, int startRow, int endRow)
        {
            base.SetCellStyleRange(cellStyleConfig, startCol, endCol, startRow, endRow);
            return this;
        }

        /// <summary>
        /// Set sheet-level global style
        /// Setup sheet-level global style
        /// </summary>
        /// <param name="styles">Style configuration function</param>
        /// <returns></returns>
        public FluentSheet SetupSheetGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            styles(_workbook, newCellStyle);
            string sheetGlobalKey = $"global_{_sheet.SheetName}";

            // Remove existing sheet global style if present
            if (_cellStylesCached.ContainsKey(sheetGlobalKey))
            {
                _cellStylesCached.Remove(sheetGlobalKey);
            }

            _cellStylesCached.Add(sheetGlobalKey, newCellStyle);
            return this;
        }


        /// <summary>
        /// Get last row with data in specified column (searches upwards from specified start row)
        /// </summary>
        /// <param name="col">Column to check</param>
        /// <param name="startRow">Start row (1-based)</param>
        /// <returns>Last row index (1-based), returns start row if not found</returns>
        private int GetLastRowWithData(ExcelCol col, int startRow)
        {
            if (_sheet == null) return startRow;

            // LastRowNum is 0-based, convert to 1-based
            int lastRowNum = _sheet.LastRowNum + 1;

            // Search upwards from last row to find first row with data
            for (int row = lastRowNum; row >= startRow; row--)
            {
                var normalizedRow = NormalizeRow(row);
                var rowObj = _sheet.GetRow(normalizedRow);

                if (rowObj != null)
                {
                    var cell = rowObj.GetCell((int)col);
                    if (cell != null && !IsCellEmpty(cell))
                    {
                        return row;
                    }
                }
            }

            // If no data row found, return start row
            return startRow;
        }

        /// <summary>
        /// Check if cell is empty
        /// </summary>
        private bool IsCellEmpty(ICell cell)
        {
            if (cell == null) return true;

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return true;
                case CellType.String:
                    return string.IsNullOrWhiteSpace(cell.StringCellValue);
                case CellType.Numeric:
                    // Can determine whether to treat 0 as empty if needed
                    return false;
                case CellType.Boolean:
                    return false;
                case CellType.Formula:
                    return false;
                default:
                    return true;
            }
        }
    }
}

