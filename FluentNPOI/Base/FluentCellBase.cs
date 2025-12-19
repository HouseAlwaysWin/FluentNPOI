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
        /// Set cell value
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="value">Value</param>
        protected void SetCellValue(ICell cell, object value)
        {
            switch (value)
            {
                case bool b:
                    cell.SetCellValue(b);
                    break;
                case DateTime dt:
                    cell.SetCellValue(dt);
                    break;
                case double d:
                    cell.SetCellValue(d);
                    break;
                case int i:
                    cell.SetCellValue((double)i);
                    break;
                case long l:
                    cell.SetCellValue((double)l);
                    break;
                case float f:
                    cell.SetCellValue((double)f);
                    break;
                case decimal m:
                    cell.SetCellValue((double)m);
                    break;
                case null:
                case DBNull:
                    cell.SetCellValue(string.Empty);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }

        /// <summary>
        /// Set cell value
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="value">Value</param>
        /// <param name="cellType">Cell Type</param>
        protected void SetCellValue(ICell cell, object value, CellType cellType)
        {
            if (cell == null) return;

            if (value == null || value == DBNull.Value)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            // 1) Handle Formula first
            if (cellType == CellType.Formula)
            {
                SetFormulaValue(cell, value);
                return;
            }

            // 2) If CellType is Unknown, just use auto-detection
            if (cellType == CellType.Unknown)
            {
                SetCellValue(cell, value);
                return;
            }

            // 3) Try to force convert based on CellType
            var text = value.ToString();
            switch (cellType)
            {
                case CellType.Boolean:
                    if (bool.TryParse(text, out var bv)) cell.SetCellValue(bv);
                    else if (int.TryParse(text, out var iv)) cell.SetCellValue(iv != 0);
                    else cell.SetCellValue(!string.IsNullOrEmpty(text));
                    break;

                case CellType.Numeric:
                    if (double.TryParse(text, out var dv)) cell.SetCellValue(dv);
                    else if (DateTime.TryParse(text, out var dtv)) cell.SetCellValue(dtv);
                    else
                    {
                        // Fallback: try setting as original value if it's already numeric
                        SetCellValue(cell, value);
                    }
                    break;

                case CellType.String:
                case CellType.Error: // NPOI error as string
                    cell.SetCellValue(text);
                    break;

                case CellType.Blank:
                    cell.SetCellValue(string.Empty);
                    break;
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
            if (cell == null) return null;

            return cell.CellType switch
            {
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue,
                CellType.String => cell.StringCellValue,
                CellType.Formula => GetCellFormulaResultValue(cell),
                CellType.Error => $"ERROR:{cell.ErrorCellValue}",
                _ => null
            };
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
            if (value == null) return default;

            if (value is T tValue) return tValue;

            try
            {
                var targetType = typeof(T);
                // Handle Nullable<T> by getting underlying type
                var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

                // DateTime conversion handling
                if (underlyingType == typeof(DateTime))
                {
                    if (value is double d)
                    {
                        try { return (T)(object)DateUtil.GetJavaDate(d); } catch { }
                    }
                    if (value is string s && DateTime.TryParse(s, out var dt))
                    {
                        return (T)(object)dt;
                    }
                }

                return (T)Convert.ChangeType(value, underlyingType);
            }
            catch
            {
                return default;
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
                for (int colIndex = startColIndex; colIndex <= endColIndex; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    ApplyCellStyleFromConfig(cell, cellStyleConfig);
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

