namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// Streaming read single row data interface
    /// </summary>
    public interface IStreamingRow
    {
        /// <summary>
        /// Row index (0-based)
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// Column count
        /// </summary>
        int ColumnCount { get; }

        /// <summary>
        /// Get value of specified column
        /// </summary>
        /// <param name="columnIndex">Column index (0-based)</param>
        /// <returns>Column value, may be null</returns>
        object GetValue(int columnIndex);

        /// <summary>
        /// Get value of specified column and convert to specified type
        /// </summary>
        /// <typeparam name="T">Target type</typeparam>
        /// <param name="columnIndex">Column index (0-based)</param>
        /// <returns>Converted value</returns>
        T GetValue<T>(int columnIndex);

        /// <summary>
        /// Check if specified column is null
        /// </summary>
        bool IsNull(int columnIndex);
    }
}
