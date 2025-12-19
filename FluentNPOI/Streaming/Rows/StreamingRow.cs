using System;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Rows
{
    /// <summary>
    /// Streaming row data implementation
    /// </summary>
    public class StreamingRow : IStreamingRow
    {
        private readonly object[] _values;

        /// <summary>
        /// Create streaming row
        /// </summary>
        /// <param name="rowIndex">Row index</param>
        /// <param name="values">Column value array</param>
        public StreamingRow(int rowIndex, object[] values)
        {
            RowIndex = rowIndex;
            _values = values ?? Array.Empty<object>();
        }

        /// <inheritdoc/>
        public int RowIndex { get; }

        /// <inheritdoc/>
        public int ColumnCount => _values.Length;

        /// <inheritdoc/>
        public object GetValue(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _values.Length)
                return null;
            return _values[columnIndex];
        }

        /// <inheritdoc/>
        public T GetValue<T>(int columnIndex)
        {
            var value = GetValue(columnIndex);
            if (value == null)
                return default;

            var targetType = typeof(T);
            var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (value is T typedValue)
                    return typedValue;

                if (value is IConvertible)
                    return (T)Convert.ChangeType(value, underlyingType);

                return default;
            }
            catch
            {
                return default;
            }
        }

        /// <inheritdoc/>
        public bool IsNull(int columnIndex)
        {
            var value = GetValue(columnIndex);
            if (value == null)
                return true;
            if (value is string str)
                return string.IsNullOrWhiteSpace(str);
            return false;
        }
    }
}
