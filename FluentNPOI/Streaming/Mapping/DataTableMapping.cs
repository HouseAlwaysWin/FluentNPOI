using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using FluentNPOI.Models;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// DataTable dedicated Mapping configuration
    /// </summary>
    public class DataTableMapping
    {
        private readonly List<ColumnMapping> _mappings = new List<ColumnMapping>();

        /// <summary>
        /// Default start row (1-based), default is 1
        /// </summary>
        public int StartRow { get; private set; } = 1;

        /// <summary>
        /// Set table default start row (1-based)
        /// </summary>
        /// <param name="row">Start row (1-based, 1st row = 1)</param>
        public DataTableMapping WithStartRow(int row)
        {
            StartRow = row < 1 ? 1 : row;
            return this;
        }

        /// <summary>
        /// Start setting column mapping
        /// </summary>
        public DataTableColumnBuilder Map(string columnName)
        {
            var mapping = new ColumnMapping { ColumnName = columnName };
            _mappings.Add(mapping);
            return new DataTableColumnBuilder(this, mapping);
        }

        /// <summary>
        /// Get all Mapping configurations
        /// </summary>
        public IReadOnlyList<ColumnMapping> GetMappings() => _mappings;

        /// <summary>
        /// Automatically generate Mapping from DataTable
        /// </summary>
        public static DataTableMapping FromDataTable(DataTable dt)
        {
            var mapping = new DataTableMapping();
            var col = ExcelCol.A;

            foreach (DataColumn column in dt.Columns)
            {
                mapping.Map(column.ColumnName)
                    .ToColumn(col)
                    .WithTitle(column.ColumnName);
                col++;
            }

            return mapping;
        }

        /// <summary>
        /// Get value from DataRow
        /// </summary>
        /// <param name="map">Column mapping configuration</param>
        /// <param name="row">DataRow data</param>
        /// <param name="rowIndex">Excel 1-based row number</param>
        /// <param name="colIndex">ExcelCol column</param>
        public object GetValue(ColumnMapping map, DataRow row, int rowIndex, ExcelCol colIndex)
        {
            if (map.ValueFunc != null)
            {
                return map.ValueFunc(row, rowIndex, colIndex);
            }

            if (!string.IsNullOrEmpty(map.ColumnName) && row.Table.Columns.Contains(map.ColumnName))
            {
                return row[map.ColumnName];
            }

            return null;
        }
    }

    /// <summary>
    /// DataTable column setting builder
    /// </summary>
    public class DataTableColumnBuilder
    {
        private readonly DataTableMapping _parent;
        private readonly ColumnMapping _mapping;

        internal DataTableColumnBuilder(DataTableMapping parent, ColumnMapping mapping)
        {
            _parent = parent;
            _mapping = mapping;
        }

        /// <summary>
        /// Set corresponding Excel column
        /// </summary>
        public DataTableColumnBuilder ToColumn(ExcelCol column)
        {
            _mapping.ColumnIndex = column;
            return this;
        }

        /// <summary>
        /// Set static title
        /// </summary>
        public DataTableColumnBuilder WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// Set dynamic title (full version)
        /// </summary>
        /// <param name="titleFunc">Title function, parameters are (row, col), returns title string</param>
        public DataTableColumnBuilder WithTitle(Func<int, ExcelCol, string> titleFunc)
        {
            _mapping.TitleFunc = titleFunc;
            return this;
        }

        /// <summary>
        /// Set static value (all rows use same value)
        /// </summary>
        /// <param name="value">Static value</param>
        public DataTableColumnBuilder WithValue(object value)
        {
            _mapping.ValueFunc = (obj, row, col) => value;
            return this;
        }

        /// <summary>
        /// Set custom value calculation (simple version, only needs DataRow)
        /// </summary>
        /// <param name="valueFunc">Value calculation function, only receives DataRow parameter</param>
        public DataTableColumnBuilder WithValue(Func<DataRow, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((DataRow)obj);
            return this;
        }

        /// <summary>
        /// Set custom value calculation (full version)
        /// </summary>
        /// <param name="valueFunc">Value calculation function, parameters are (row, excelRow, col), excelRow is Excel 1-based row number, col is ExcelCol column</param>
        public DataTableColumnBuilder WithValue(Func<DataRow, int, ExcelCol, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((DataRow)obj, row, col);
            return this;
        }

        /// <summary>
        /// Set static formula (simple version)
        /// </summary>
        /// <param name="formula">Formula string (without =)</param>
        public DataTableColumnBuilder WithFormula(string formula)
        {
            _mapping.FormulaFunc = (row, col) => formula;
            return this;
        }

        /// <summary>
        /// Set dynamic formula (full version)
        /// </summary>
        /// <param name="formulaFunc">Formula function, parameters are (row, col), returns formula string (without =)</param>
        public DataTableColumnBuilder WithFormula(Func<int, ExcelCol, string> formulaFunc)
        {
            _mapping.FormulaFunc = formulaFunc;
            return this;
        }

        /// <summary>
        /// Set data style
        /// </summary>
        public DataTableColumnBuilder WithStyle(string styleKey)
        {
            _mapping.StyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// Set title style
        /// </summary>
        public DataTableColumnBuilder WithTitleStyle(string styleKey)
        {
            _mapping.TitleStyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// Set cell type
        /// </summary>
        public DataTableColumnBuilder WithCellType(NPOI.SS.UserModel.CellType cellType)
        {
            _mapping.CellType = cellType;
            return this;
        }

        /// <summary>
        /// Copy title style from specified cell
        /// </summary>
        /// <param name="row">Source row (1-based, 1st row = 1)</param>
        /// <param name="col">Source column</param>
        public DataTableColumnBuilder WithTitleStyleFrom(int row, ExcelCol col)
        {
            _mapping.TitleStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// Copy data style from specified cell
        /// </summary>
        /// <param name="row">Source row (1-based, 1st row = 1)</param>
        /// <param name="col">Source column</param>
        public DataTableColumnBuilder WithStyleFrom(int row, ExcelCol col)
        {
            _mapping.DataStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// Set column row offset (offset downwards from table start row)
        /// </summary>
        /// <param name="offset">Offset amount (positive number means downward offset, default 0)</param>
        public DataTableColumnBuilder WithRowOffset(int offset)
        {
            _mapping.RowOffset = offset;
            return this;
        }

        /// <summary>
        /// Continue to set next column
        /// </summary>
        public DataTableColumnBuilder Map(string columnName)
        {
            return _parent.Map(columnName);
        }
    }
}
