using NPOI.SS.UserModel;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Base
{
    /// <summary>
    /// Base cell operation class
    /// </summary>
    public abstract class FluentCellBase : FluentWorkbookBase
    {
        protected ISheet _sheet;

        protected FluentCellBase(IWorkbook workbook, ISheet sheet, Dictionary<string, ICellStyle> cellStylesCached)
            : base(workbook, cellStylesCached)
        {
            _sheet = sheet;
        }



        /// <summary>
        /// Set cell style
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="cellNameMap">Cell name mapping</param>
        /// <param name="cellStyleParams">Cell style parameters</param>
        protected void SetCellStyle(ICell cell, TableCellSet cellNameMap, TableCellStyleParams cellStyleParams)
        {
            // If there is a dynamic style setting function, use it first
            if (cellNameMap.SetCellStyleAction != null)
            {
                // ✅ Call function to get style configuration first
                CellStyleConfig config = cellNameMap.SetCellStyleAction(cellStyleParams);
                ApplyCellStyleFromConfig(cell, config);
            }
            // If there is a fixed style key, use cached style
            else if (!string.IsNullOrWhiteSpace(cellNameMap.CellStyleKey) && _cellStylesCached.ContainsKey(cellNameMap.CellStyleKey))
            {
                cell.CellStyle = _cellStylesCached[cellNameMap.CellStyleKey];
            }
            // If neither, check sheet-level global style first, then workbook-level global style
            // If none specified, check sheet-level global style first, then workbook-level global style
            else
            {
                string sheetGlobalKey = $"global_{_sheet.SheetName}";
                if (_cellStylesCached.ContainsKey(sheetGlobalKey))
                {
                    cell.CellStyle = _cellStylesCached[sheetGlobalKey];
                }
                else if (_cellStylesCached.ContainsKey("global"))
                {
                    cell.CellStyle = _cellStylesCached["global"];
                }
            }
        }

        /// <summary>
        /// Set cell value
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="value">Value</param>
        protected void SetCellValue(ICell cell, object value)
        {
            if (value is bool b)
            {
                cell.SetCellValue(b);
                return;
            }
            if (value is DateTime dt)
            {
                cell.SetCellValue(dt);
                return;
            }
            if (value is int i)
            {
                cell.SetCellValue((double)i);
                return;
            }
            if (value is long l)
            {
                cell.SetCellValue((double)l);
                return;
            }
            if (value is float f)
            {
                cell.SetCellValue((double)f);
                return;
            }
            if (value is double d)
            {
                cell.SetCellValue(d);
                return;
            }
            if (value is decimal m)
            {
                cell.SetCellValue((double)m);
                return;
            }

            cell.SetCellValue(value.ToString());
        }

        /// <summary>
        /// Set cell value
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="value">Value</param>
        /// <param name="cellType">Cell Type</param>
        protected void SetCellValue(ICell cell, object value, CellType cellType)
        {
            if (cell == null)
                return;

            if (value == null || value == DBNull.Value)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            // 1) Write based on actual type of value first
            SetCellValue(cell, value);

            // 2) If CellType is specified (and not Unknown), override with CellType
            if (cellType == CellType.Unknown) return;
            if (cellType == CellType.Formula)
            {
                SetFormulaValue(cell, value);
                return;
            }

            var text = value.ToString();
            switch (cellType)
            {
                case CellType.Boolean:
                    {
                        if (bool.TryParse(text, out var bv)) { cell.SetCellValue(bv); return; }
                        if (int.TryParse(text, out var iv)) { cell.SetCellValue(iv != 0); return; }
                        cell.SetCellValue(!string.IsNullOrEmpty(text));
                        return;
                    }
                case CellType.Numeric:
                    {
                        if (double.TryParse(text, out var dv)) { cell.SetCellValue(dv); return; }
                        if (DateTime.TryParse(text, out var dtv)) { cell.SetCellValue(dtv); return; }
                        // If unable to convert to numeric/date, keep previous write result
                        return;
                    }
                case CellType.String:
                    {
                        cell.SetCellValue(text);
                        return;
                    }
                case CellType.Blank:
                    {
                        cell.SetCellValue(string.Empty);
                        return;
                    }
                case CellType.Error:
                    {
                        // NPOI error type cannot be set directly from object, fallback to string display
                        cell.SetCellValue(text);
                        return;
                    }
                default:
                    return;
            }
        }

        /// <summary>
        /// Set formula value
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="value">Value</param>
        protected void SetFormulaValue(ICell cell, object value)
        {
            if (cell == null) return;
            if (value == null || value == DBNull.Value) return;

            var formula = value.ToString();
            if (string.IsNullOrWhiteSpace(formula)) return;

            // NPOI SetCellFormula requires pure formula string (without '=')
            if (formula.StartsWith("=")) formula = formula.Substring(1);

            cell.SetCellFormula(formula);
        }

        /// <summary>
        /// Get cell value, return corresponding C# type based on cell type
        /// </summary>
        /// <param name="cell">Cell to read</param>
        /// <returns>Cell value (bool, DateTime, double, string or null)</returns>
        protected object GetCellValue(ICell cell)
        {
            if (cell == null)
                return null;

            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Numeric:
                    // Check if date format
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue;

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Formula:
                    // For formula, return calculated value
                    return GetCellFormulaResultValue(cell);

                case CellType.Blank:
                    return null;

                case CellType.Error:
                    return $"ERROR:{cell.ErrorCellValue}";

                default:
                    return null;
            }
        }

        /// <summary>
        /// Get cell value and convert to specific type
        /// </summary>
        /// <typeparam name="T">Target Type</typeparam>
        /// <param name="cell">Cell to read</param>
        /// <returns>Converted value</returns>
        protected T GetCellValue<T>(ICell cell)
        {
            var value = GetCellValue(cell);

            if (value == null)
                return default(T);

            try
            {
                // If type matches already, return directly
                if (value is T result)
                    return result;

                // Special handling for DateTime
                if (typeof(T) == typeof(DateTime) || typeof(T) == typeof(DateTime?))
                {
                    if (value is DateTime dt)
                        return (T)(object)dt;

                    // If value is double (Excel stores dates as numbers), try to convert
                    if (value is double d && cell != null)
                    {
                        // Check if cell has date format
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return (T)(object)cell.DateCellValue;
                        }
                        // Try to convert number to date (OLE Automation Date)
                        try
                        {
                            var dateTime = DateUtil.GetJavaDate(d);
                            return (T)(object)dateTime;
                        }
                        catch
                        {
                            // If conversion fails, return default
                        }
                    }

                    // Try string conversion
                    if (value is string str && DateTime.TryParse(str, out var parsedDate))
                    {
                        return (T)(object)parsedDate;
                    }

                    return default(T);
                }

                // Try conversion
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch
            {
                return default(T);
            }
        }

        /// <summary>
        /// Get cell formula string
        /// </summary>
        /// <param name="cell">Cell to read</param>
        /// <returns>Formula string (without '=' prefix), or null if not a formula cell</returns>
        protected string GetCellFormulaValue(ICell cell)
        {
            if (cell == null)
                return null;

            if (cell.CellType == CellType.Formula)
            {
                return cell.CellFormula;
            }

            return null;
        }

        /// <summary>
        /// Get calculated result value of formula cell
        /// </summary>
        /// <param name="cell">Formula cell</param>
        /// <returns>Calculated value</returns>
        private object GetCellFormulaResultValue(ICell cell)
        {
            if (cell == null || cell.CellType != CellType.Formula)
                return null;

            try
            {
                switch (cell.CachedFormulaResultType)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return cell.DateCellValue;
                        }
                        return cell.NumericCellValue;

                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Blank:
                        return null;

                    case CellType.Error:
                        return $"ERROR:{cell.ErrorCellValue}";

                    default:
                        return null;
                }
            }
            catch
            {
                // If unable to get calculated result, return null
                return null;
            }
        }


        /// <summary>
        /// Apply cell style from configuration (check cache, create and cache if not exists)
        /// </summary>
        /// <param name="cell">Cell to apply style to</param>
        /// <param name="config">Style configuration</param>
        private void ApplyCellStyleFromConfig(ICell cell, CellStyleConfig config)
        {
            if (!string.IsNullOrWhiteSpace(config.Key))
            {
                // ✅ Check cache first
                if (!_cellStylesCached.ContainsKey(config.Key))
                {
                    // ✅ Create new style only if not exists
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    config.StyleSetter(newCellStyle);
                    _cellStylesCached.Add(config.Key, newCellStyle);
                }
                // Always use cached style
                cell.CellStyle = _cellStylesCached[config.Key];
            }
            else
            {
                throw new Exception("cell style key can not be null or empty");
            }
        }

        /// <summary>
        /// Set style for a range of cells
        /// </summary>
        /// <param name="cellStyleConfig">Style configuration</param>
        /// <param name="startCol">Start column</param>
        /// <param name="endCol">End column</param>
        /// <param name="startRow">Start row</param>
        /// <param name="endRow">End row</param>   
        protected void SetCellStyleRange(
            CellStyleConfig cellStyleConfig,
             ExcelCol startCol, ExcelCol endCol, int startRow, int endRow)
        {
            int startColIndex = (int)startCol;
            int endColIndex = (int)endCol;
            int startRowIndex = NormalizeRow(startRow);
            int endRowIndex = NormalizeRow(endRow);

            for (int rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
            {
                IRow row = _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);
                if (row != null)
                {
                    for (int colIndex = startColIndex; colIndex <= endColIndex; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                        ApplyCellStyleFromConfig(cell, cellStyleConfig);
                    }

                }
            }
        }

        protected ICell SetCellPositionInternal(ExcelCol col, int row)
        {
            if (_sheet == null) throw new System.InvalidOperationException("No active sheet. Call UseSheet(...) first.");

            var normalizedRow = NormalizeRow(row);

            var rowObj = _sheet.GetRow(normalizedRow) ?? _sheet.CreateRow(normalizedRow);
            var cell = rowObj.GetCell((int)col) ?? rowObj.CreateCell((int)col);
            return cell;
        }
    }
}

