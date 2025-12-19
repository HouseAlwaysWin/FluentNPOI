using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.HSSF.UserModel; // Added for .xls support
using FluentNPOI.Streaming;
using FluentNPOI.Streaming.Mapping;
using System;
using System.IO;
using System.Collections.Generic;
using FluentNPOI.Models;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// Builder for Pipeline Processing (Read-Modify-Write)
    /// Supports both Streaming (SXSSF for .xlsx) and DOM (HSSF for .xls) backends.
    /// </summary>
    /// <typeparam name="T">Data Model Type</typeparam>
    /// <example>
    /// <code>
    /// FluentWorkbook.Stream&lt;MyData&gt;("input.xlsx")
    ///     .Transform(x => x.Value *= 2)
    ///     .SaveAs("output.xlsx");
    /// </code>
    /// </example>
    public class StreamingBuilder<T> where T : new()
    {
        private readonly Stream _inputStream;
        private readonly bool _ownsStream;
        private Action<T> _transform;
        private Action<FluentSheet> _configureSheet;
        private FluentMapping<T> _mapping;
        private string _sheetName;

        internal StreamingBuilder(string inputFile)
        {
            if (string.IsNullOrEmpty(inputFile)) throw new ArgumentNullException(nameof(inputFile));
            _inputStream = File.OpenRead(inputFile);
            _ownsStream = true;
        }

        internal StreamingBuilder(Stream inputStream)
        {
            _inputStream = inputStream ?? throw new ArgumentNullException(nameof(inputStream));
            _ownsStream = false;
        }

        /// <summary>
        /// Set transformation logic (executed per row)
        /// </summary>
        public StreamingBuilder<T> Transform(Action<T> transform)
        {
            _transform = transform;
            return this;
        }

        /// <summary>
        /// Configure Sheet settings (Styles, Widths, etc.)
        /// </summary>
        public StreamingBuilder<T> Configure(Action<FluentSheet> configure)
        {
            _configureSheet = configure;
            return this;
        }

        /// <summary>
        /// Use custom FluentMapping
        /// </summary>
        public StreamingBuilder<T> WithMapping(FluentMapping<T> mapping)
        {
            _mapping = mapping;
            return this;
        }

        /// <summary>
        /// Filter input sheet name
        /// </summary>
        public StreamingBuilder<T> UseSheet(string sheetName)
        {
            _sheetName = sheetName;
            return this;
        }

        /// <summary>
        /// Execute Pipeline and Save to File
        /// </summary>
        public void SaveAs(string outputFile)
        {
            if (string.IsNullOrEmpty(outputFile)) throw new ArgumentNullException(nameof(outputFile));

            var dir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // Detect format
            bool useDom = outputFile.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);

            using (var outStream = File.Create(outputFile))
            {
                if (useDom)
                {
                    SaveToDom(outStream); // .xls (HSSF)
                }
                else
                {
                    SaveTo(outStream); // .xlsx (SXSSF - Streaming)
                }
            }
        }

        /// <summary>
        /// Execute Pipeline and Save to Stream (Defaults to XLSX Streaming)
        /// </summary>
        public void SaveTo(Stream outputStream)
        {
            try
            {
                // 1. Initialize Streaming Writer (SXSSF)
                using (var workbook = new SXSSFWorkbook(100))
                {
                    ProcessWorkbook(workbook, outputStream);
                }
            }
            finally
            {
                if (_ownsStream)
                {
                    _inputStream?.Dispose();
                }
            }
        }

        /// <summary>
        /// Execute Pipeline and Save to Stream using DOM (HSSF for .xls)
        /// Note: This is NOT streaming low-memory, but provides same API pipeline for legacy files.
        /// </summary>
        private void SaveToDom(Stream outputStream)
        {
            try
            {
                // 1. Initialize DOM Writer (HSSF)
                using (var workbook = new HSSFWorkbook())
                {
                    ProcessWorkbook(workbook, outputStream);
                }
            }
            finally
            {
                if (_ownsStream)
                {
                    _inputStream?.Dispose();
                }
            }
        }

        private void ProcessWorkbook(IWorkbook workbook, Stream outputStream)
        {
            var sheet = workbook.CreateSheet(_sheetName ?? "Sheet1");
            var styleCache = new Dictionary<string, ICellStyle>();

            // 2. Configure Sheet (Metadata)
            var fluentSheet = new FluentSheet(workbook, sheet, styleCache);
            _configureSheet?.Invoke(fluentSheet);

            // 3. Prepare Mapping
            var mapping = _mapping ?? new FluentMapping<T>();

            // 4. Read & Loop
            // Note: ExcelDataReader works for both .xls and .xlsx input transparently
            var sourceData = FluentExcelReader.Read<T>(_inputStream, _sheetName);

            IEnumerable<T> Pipeline()
            {
                foreach (var item in sourceData)
                {
                    _transform?.Invoke(item);
                    yield return item;
                }
            }

            // 5. Write Table
            fluentSheet.SetTable(Pipeline(), mapping)
                       .BuildRows();

            // 6. Write Output
            workbook.Write(outputStream);
        }
    }
}
