using System.Collections.Generic;
using FluentNPOI.Streaming.Abstractions;
using FluentNPOI.Streaming.Mapping;
using FluentNPOI.Streaming.Pipeline;
using FluentNPOI.Streaming.Readers;

namespace FluentNPOI.Streaming
{
    /// <summary>
    /// FluentNPOI streaming reader entry point
    /// </summary>
    public static class FluentExcelReader
    {
        /// <summary>
        /// Read Excel using Header auto-mapping
        /// </summary>
        /// <typeparam name="T">Target type</typeparam>
        /// <param name="filePath">File path</param>
        /// <param name="sheetName">Sheet name (optional)</param>
        /// <returns>Object enumeration</returns>
        public static IEnumerable<T> Read<T>(string filePath, string sheetName = null) where T : new()
        {
            using (var reader = new ExcelDataReaderAdapter(filePath))
            {
                if (!string.IsNullOrEmpty(sheetName))
                    reader.SelectSheet(sheetName);

                // Read Header to create auto Mapping
                var headers = reader.ReadHeader();
                var mapper = CreateAutoMapper<T>(headers);

                foreach (var row in reader.ReadRows())
                {
                    yield return mapper.Map(row);
                }
            }
        }

        /// <summary>
        /// Read Excel using FluentMapping
        /// </summary>
        /// <typeparam name="T">Target type</typeparam>
        /// <param name="filePath">File path</param>
        /// <param name="mapping">Mapping configuration</param>
        /// <param name="sheetName">Sheet name (optional)</param>
        /// <returns>Object enumeration</returns>
        public static IEnumerable<T> Read<T>(string filePath, FluentMapping<T> mapping, string sheetName = null) where T : new()
        {
            using (var reader = new ExcelDataReaderAdapter(filePath))
            {
                if (!string.IsNullOrEmpty(sheetName))
                    reader.SelectSheet(sheetName);

                var pipeline = StreamingPipelineBuilder.CreatePipeline(reader, mapping)
                    .SkipHeader();

                foreach (var item in pipeline.ToEnumerable())
                {
                    yield return item;
                }
            }
        }

        /// <summary>
        /// Create pipeline for more granular control
        /// </summary>
        public static StreamingPipeline<T> CreatePipeline<T>(string filePath, FluentMapping<T> mapping) where T : new()
        {
            var reader = new ExcelDataReaderAdapter(filePath);
            return StreamingPipelineBuilder.CreatePipeline(reader, mapping);
        }

        private static IRowMapper<T> CreateAutoMapper<T>(string[] headers) where T : new()
        {
            var mapping = new FluentMapping<T>();
            var properties = typeof(T).GetProperties();

            for (int i = 0; i < headers.Length; i++)
            {
                var headerName = headers[i]?.Trim();
                if (string.IsNullOrEmpty(headerName))
                    continue;

                // Find property with matching name
                foreach (var prop in properties)
                {
                    if (string.Equals(prop.Name, headerName, System.StringComparison.OrdinalIgnoreCase))
                    {
                        // Use reflection to set mapping (simplified version)
                        mapping.AddInternalMapping(prop, (FluentNPOI.Models.ExcelCol)i);
                        break;
                    }
                }
            }

            return mapping;
        }
    }
}
