using System;
using System.Collections.Generic;

namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// Streaming Excel reader interface
    /// </summary>
    public interface IStreamingReader : IDisposable
    {
        /// <summary>
        /// Get all sheet names
        /// </summary>
        IReadOnlyList<string> SheetNames { get; }

        /// <summary>
        /// Select sheet (by name)
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>Is success</returns>
        bool SelectSheet(string sheetName);

        /// <summary>
        /// Select sheet (by index)
        /// </summary>
        /// <param name="sheetIndex">Sheet index (0-based)</param>
        /// <returns>Is success</returns>
        bool SelectSheet(int sheetIndex);

        /// <summary>
        /// Stream read all rows
        /// </summary>
        /// <returns>Row enumerator</returns>
        IEnumerable<IStreamingRow> ReadRows();

        /// <summary>
        /// Read Header row (first row)
        /// </summary>
        /// <returns>Header name array</returns>
        string[] ReadHeader();
    }
}
