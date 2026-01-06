using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using FluentNPOI.Models;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// Fluent Mapping configuration, used to define mapping between Excel columns and properties
    /// </summary>
    /// <typeparam name="T">Target type</typeparam>
    public class FluentMapping<T> : IRowMapper<T> where T : new()
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
        public FluentMapping<T> WithStartRow(int row)
        {
            StartRow = row < 1 ? 1 : row;
            return this;
        }

        /// <summary>
        /// Start setting property mapping
        /// </summary>
        public FluentColumnBuilder<T> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            var propertyInfo = GetPropertyInfo(expression);
            var mapping = new ColumnMapping { Property = propertyInfo };
            _mappings.Add(mapping);
            return new FluentColumnBuilder<T>(this, mapping);
        }

        /// <summary>
        /// Get all Mapping configurations
        /// </summary>
        public IReadOnlyList<ColumnMapping> GetMappings() => _mappings;

        /// <summary>
        /// Add Mapping directly (for auto mapping from column headers)
        /// </summary>
        public void AddInternalMapping(PropertyInfo property, ExcelCol column)
        {
            var mapping = new ColumnMapping
            {
                Property = property,
                ColumnIndex = column
            };
            _mappings.Add(mapping);
        }

        /// <summary>
        /// Convert streaming row to DTO (IRowMapper implementation)
        /// </summary>
        public T Map(IStreamingRow row)
        {
            var instance = new T();

            foreach (var mapping in _mappings)
            {
                if (!mapping.ColumnIndex.HasValue)
                    continue;

                var value = row.GetValue((int)mapping.ColumnIndex.Value);
                if (value == null)
                    continue;

                try
                {
                    var targetType = mapping.Property.PropertyType;
                    var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

                    object convertedValue;
                    if (value.GetType() == underlyingType)
                    {
                        convertedValue = value;
                    }
                    else if (value is IConvertible)
                    {
                        convertedValue = Convert.ChangeType(value, underlyingType);
                    }
                    else
                    {
                        continue;
                    }

                    mapping.Property.SetValue(instance, convertedValue);
                }
                catch
                {
                    // Conversion failed, skip
                }
            }

            return instance;
        }

        private PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            if (expression.Body is MemberExpression member)
                return member.Member as PropertyInfo;
            if (expression.Body is UnaryExpression unary && unary.Operand is MemberExpression unaryMember)
                return unaryMember.Member as PropertyInfo;
            throw new ArgumentException("Expression must be a property selector");
        }
    }

    /// <summary>
    /// Fluent column setting builder
    /// </summary>
    public class FluentColumnBuilder<T> where T : new()
    {
        private readonly FluentMapping<T> _parent;
        private readonly ColumnMapping _mapping;

        internal FluentColumnBuilder(FluentMapping<T> parent, ColumnMapping mapping)
        {
            _parent = parent;
            _mapping = mapping;
        }

        /// <summary>
        /// Set corresponding Excel column
        /// </summary>
        public FluentColumnBuilder<T> ToColumn(ExcelCol column)
        {
            _mapping.ColumnIndex = column;
            return this;
        }

        /// <summary>
        /// Set static title (used when writing)
        /// </summary>
        public FluentColumnBuilder<T> WithTitle(string title)
        {
            _mapping.Title = title;
            return this;
        }

        /// <summary>
        /// Set dynamic title (full version) (used when writing)
        /// </summary>
        /// <param name="titleFunc">Title function, parameters are (row, col), returns title string</param>
        public FluentColumnBuilder<T> WithTitle(Func<int, ExcelCol, string> titleFunc)
        {
            _mapping.TitleFunc = titleFunc;
            return this;
        }

        /// <summary>
        /// Set format (used when writing)
        /// </summary>
        public FluentColumnBuilder<T> WithFormat(string format)
        {
            _mapping.Format = format;
            return this;
        }

        /// <summary>
        /// Set static value (all rows use same value) (used when writing)
        /// </summary>
        /// <param name="value">Static value</param>
        public FluentColumnBuilder<T> WithValue(object value)
        {
            _mapping.ValueFunc = (obj, row, col) => value;
            return this;
        }

        /// <summary>
        /// Set custom value calculation (simple version, only needs data object) (used when writing)
        /// </summary>
        /// <param name="valueFunc">Value calculation function, only receives data object parameter</param>
        public FluentColumnBuilder<T> WithValue(Func<T, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((T)obj);
            return this;
        }

        /// <summary>
        /// Set custom value calculation (full version) (used when writing)
        /// </summary>
        /// <param name="valueFunc">Value calculation function, parameters are (item, row, col), row is Excel 1-based row number, col is ExcelCol column</param>
        public FluentColumnBuilder<T> WithValue(Func<T, int, ExcelCol, object> valueFunc)
        {
            _mapping.ValueFunc = (obj, row, col) => valueFunc((T)obj, row, col);
            return this;
        }

        /// <summary>
        /// Set static formula (simple version) (used when writing)
        /// </summary>
        /// <param name="formula">Formula string (without =)</param>
        public FluentColumnBuilder<T> WithFormula(string formula)
        {
            _mapping.FormulaFunc = (row, col) => formula;
            return this;
        }

        /// <summary>
        /// Set dynamic formula (full version) (used when writing)
        /// </summary>
        /// <param name="formulaFunc">Formula function, parameters are (row, col), returns formula string (without =)</param>
        public FluentColumnBuilder<T> WithFormula(Func<int, ExcelCol, string> formulaFunc)
        {
            _mapping.FormulaFunc = formulaFunc;
            return this;
        }

        /// <summary>
        /// Set data cell style (use style Key)
        /// </summary>
        public FluentColumnBuilder<T> WithStyle(string styleKey)
        {
            _mapping.StyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// Set title cell style (use style Key)
        /// </summary>
        public FluentColumnBuilder<T> WithTitleStyle(string styleKey)
        {
            _mapping.TitleStyleKey = styleKey;
            _mapping.TitleStyleKey = styleKey;
            return this;
        }

        /// <summary>
        /// Copy title style from specified cell
        /// </summary>
        /// <param name="row">Source row (1-based, 1st row = 1)</param>
        /// <param name="col">Source column</param>
        public FluentColumnBuilder<T> WithTitleStyleFrom(int row, ExcelCol col)
        {
            _mapping.TitleStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// Copy data style from specified cell
        /// </summary>
        /// <param name="row">Source row (1-based, 1st row = 1)</param>
        /// <param name="col">Source column</param>
        public FluentColumnBuilder<T> WithStyleFrom(int row, ExcelCol col)
        {
            _mapping.DataStyleRef = StyleReference.FromUserInput(row, col);
            return this;
        }

        /// <summary>
        /// Set cell type
        /// </summary>
        public FluentColumnBuilder<T> WithCellType(NPOI.SS.UserModel.CellType cellType)
        {
            _mapping.CellType = cellType;
            return this;
        }

        /// <summary>
        /// Set dynamic style (determine style Key based on data)
        /// </summary>
        /// <param name="styleFunc">Receives data object, returns style Key</param>
        public FluentColumnBuilder<T> WithDynamicStyle(Func<T, string> styleFunc)
        {
            _mapping.DynamicStyleFunc = obj => styleFunc((T)obj);
            return this;
        }

        /// <summary>
        /// Set column row offset (offset downwards from table start row)
        /// </summary>
        /// <param name="offset">Offset amount (positive number means downward offset, default 0)</param>
        public FluentColumnBuilder<T> WithRowOffset(int offset)
        {
            _mapping.RowOffset = offset;
            return this;
        }

        /// <summary>
        /// Set other style configurations (will be merged into automatically generated style)
        /// </summary>
        public FluentColumnBuilder<T> WithStyleConfig(Action<NPOI.SS.UserModel.IWorkbook, NPOI.SS.UserModel.ICellStyle> config)
        {
            EnsureStyleKey();
            _mapping.StyleConfig += config;
            return this;
        }

        /// <summary>
        /// Set number format
        /// </summary>
        public FluentColumnBuilder<T> WithNumberFormat(string format)
        {
            return WithStyleConfig((wb, style) =>
            {
                var df = wb.CreateDataFormat();
                style.DataFormat = df.GetFormat(format);
            });
        }

        /// <summary>
        /// Set background color
        /// </summary>
        public FluentColumnBuilder<T> WithBackgroundColor(NPOI.SS.UserModel.IndexedColors color)
        {
            return WithStyleConfig((wb, style) =>
            {
                style.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                style.FillForegroundColor = color.Index;
            });
        }

        /// <summary>
        /// Set font
        /// </summary>
        public FluentColumnBuilder<T> WithFont(string fontName = null, double? fontSize = null, bool isBold = false)
        {
            return WithStyleConfig((wb, style) =>
            {
                var font = wb.CreateFont();
                if (fontName != null) font.FontName = fontName;
                if (fontSize.HasValue) font.FontHeightInPoints = fontSize.Value;
                font.IsBold = isBold;
                style.SetFont(font);
            });
        }

        /// <summary>
        /// Set border
        /// </summary>
        public FluentColumnBuilder<T> WithBorder(NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            return WithStyleConfig((wb, style) =>
            {
                style.BorderTop = borderStyle;
                style.BorderBottom = borderStyle;
                style.BorderLeft = borderStyle;
                style.BorderRight = borderStyle;
            });
        }

        /// <summary>
        /// Set alignment
        /// </summary>
        public FluentColumnBuilder<T> WithAlignment(NPOI.SS.UserModel.HorizontalAlignment horizontal = NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.VerticalAlignment vertical = NPOI.SS.UserModel.VerticalAlignment.Center)
        {
            return WithStyleConfig((wb, style) =>
            {
                style.Alignment = horizontal;
                style.VerticalAlignment = vertical;
            });
        }

        public FluentColumnBuilder<T> WithWrapText(bool wrap = true)
        {
            return WithStyleConfig((wb, style) =>
            {
                style.WrapText = wrap;
            });
        }

        /// <summary>
        /// Set title style configuration (will be merged into automatically generated style)
        /// </summary>
        public FluentColumnBuilder<T> WithTitleStyleConfig(Action<NPOI.SS.UserModel.IWorkbook, NPOI.SS.UserModel.ICellStyle> config)
        {
            EnsureTitleStyleKey();
            _mapping.TitleStyleConfig += config;
            return this;
        }

        /// <summary>
        /// Set title font
        /// </summary>
        public FluentColumnBuilder<T> WithTitleFont(string fontName = null, double? fontSize = null, bool isBold = false, NPOI.SS.UserModel.IndexedColors color = null)
        {
            return WithTitleStyleConfig((wb, style) =>
            {
                var font = wb.CreateFont();
                if (fontName != null) font.FontName = fontName;
                if (fontSize.HasValue) font.FontHeightInPoints = fontSize.Value;
                if (color != null) font.Color = color.Index;
                font.IsBold = isBold;
                style.SetFont(font);
            });
        }

        /// <summary>
        /// Set title background color
        /// </summary>
        public FluentColumnBuilder<T> WithTitleBackgroundColor(NPOI.SS.UserModel.IndexedColors color)
        {
            return WithTitleStyleConfig((wb, style) =>
            {
                style.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                style.FillForegroundColor = color.Index;
            });
        }

        /// <summary>
        /// Set title alignment
        /// </summary>
        public FluentColumnBuilder<T> WithTitleAlignment(NPOI.SS.UserModel.HorizontalAlignment horizontal = NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.VerticalAlignment vertical = NPOI.SS.UserModel.VerticalAlignment.Center)
        {
            return WithTitleStyleConfig((wb, style) =>
            {
                style.Alignment = horizontal;
                style.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        /// Set title border
        /// </summary>
        public FluentColumnBuilder<T> WithTitleBorder(NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            return WithTitleStyleConfig((wb, style) =>
            {
                style.BorderTop = borderStyle;
                style.BorderBottom = borderStyle;
                style.BorderLeft = borderStyle;
                style.BorderRight = borderStyle;
            });
        }

        private void EnsureStyleKey()
        {
            if (string.IsNullOrEmpty(_mapping.GeneratedStyleKey))
            {
                _mapping.GeneratedStyleKey = $"AutoStyle_{Guid.NewGuid()}";
            }
        }

        private void EnsureTitleStyleKey()
        {
            if (string.IsNullOrEmpty(_mapping.GeneratedTitleStyleKey))
            {
                _mapping.GeneratedTitleStyleKey = $"AutoTitleStyle_{Guid.NewGuid()}";
            }
        }

        /// <summary>
        /// Continue to set next property
        /// </summary>
        public FluentColumnBuilder<T> Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            return _parent.Map(expression);
        }
    }

    /// <summary>
    /// Column mapping configuration
    /// </summary>
    public class ColumnMapping
    {
        public PropertyInfo Property { get; set; }
        public ExcelCol? ColumnIndex { get; set; }
        public string Title { get; set; }
        public Func<int, ExcelCol, string> TitleFunc { get; set; }
        public string Format { get; set; }

        // Used when writing
        public Func<object, int, ExcelCol, object> ValueFunc { get; set; }
        public Func<int, ExcelCol, string> FormulaFunc { get; set; }
        public string StyleKey { get; set; }
        public string TitleStyleKey { get; set; }
        public NPOI.SS.UserModel.CellType? CellType { get; set; }
        public Func<object, string> DynamicStyleFunc { get; set; }

        /// <summary>
        /// Static style configuration (used for dynamic style generation)
        /// </summary>
        public Action<NPOI.SS.UserModel.IWorkbook, NPOI.SS.UserModel.ICellStyle> StyleConfig { get; set; }

        /// <summary>
        /// Title style configuration (used for dynamic style generation)
        /// </summary>
        public Action<NPOI.SS.UserModel.IWorkbook, NPOI.SS.UserModel.ICellStyle> TitleStyleConfig { get; set; }

        /// <summary>
        /// Automatically generated style Key (internal use)
        /// </summary>
        public string GeneratedStyleKey { get; set; }

        /// <summary>
        /// Automatically generated title style Key (internal use)
        /// </summary>
        public string GeneratedTitleStyleKey { get; set; }

        // Column name (for DataTable)
        public string ColumnName { get; set; }

        // Style reference
        public StyleReference TitleStyleRef { get; set; }
        public StyleReference DataStyleRef { get; set; }

        /// <summary>
        /// Column row offset (default 0, positive number means downward offset)
        /// </summary>
        public int RowOffset { get; set; } = 0;
    }

    public class StyleReference
    {
        public int Row { get; set; }
        public ExcelCol Column { get; set; }

        /// <summary>
        /// Create StyleReference from user input (automatically convert 1-based row to 0-based)
        /// </summary>
        /// <param name="row">User input row number (1-based, 1st row = 1)</param>
        /// <param name="col">Column</param>
        public static StyleReference FromUserInput(int row, ExcelCol col)
        {
            return new StyleReference
            {
                Row = row < 1 ? 0 : row - 1,  // Convert 1-based to 0-based
                Column = col
            };
        }
    }
}
