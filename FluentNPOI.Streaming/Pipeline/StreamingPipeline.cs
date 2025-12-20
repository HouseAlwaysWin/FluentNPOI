using System;
using System.Collections.Generic;
using FluentNPOI.Streaming.Abstractions;

namespace FluentNPOI.Streaming.Pipeline
{
    /// <summary>
    /// Streaming processing pipeline, providing Fluent chaining API
    /// </summary>
    /// <typeparam name="T">Target type</typeparam>
    public class StreamingPipeline<T> where T : new()
    {
        private readonly IStreamingReader _reader;
        private readonly IRowMapper<T> _mapper;
        private int _skipRows;
        private Func<IStreamingRow, bool> _filter;

        internal StreamingPipeline(IStreamingReader reader, IRowMapper<T> mapper)
        {
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _mapper = mapper ?? throw new ArgumentNullException(nameof(mapper));
        }

        /// <summary>
        /// Skip specified number of rows (e.g. Header)
        /// </summary>
        public StreamingPipeline<T> Skip(int rowCount)
        {
            _skipRows = rowCount;
            return this;
        }

        /// <summary>
        /// Skip Header (equivalent to Skip(1))
        /// </summary>
        public StreamingPipeline<T> SkipHeader()
        {
            return Skip(1);
        }

        /// <summary>
        /// Filter rows
        /// </summary>
        public StreamingPipeline<T> Where(Func<IStreamingRow, bool> predicate)
        {
            _filter = predicate;
            return this;
        }

        /// <summary>
        /// Execute pipeline and return result (deferred execution)
        /// </summary>
        public IEnumerable<T> ToEnumerable()
        {
            int skipped = 0;

            foreach (var row in _reader.ReadRows())
            {
                // Skip specified number of rows
                if (skipped < _skipRows)
                {
                    skipped++;
                    continue;
                }

                // Apply filter
                if (_filter != null && !_filter(row))
                    continue;

                // Map and return
                yield return _mapper.Map(row);
            }
        }

        /// <summary>
        /// Execute pipeline and return List
        /// </summary>
        public List<T> ToList()
        {
            var result = new List<T>();
            foreach (var item in ToEnumerable())
            {
                result.Add(item);
            }
            return result;
        }
    }

    /// <summary>
    /// Pipeline builder
    /// </summary>
    public static class StreamingPipelineBuilder
    {
        /// <summary>
        /// Create pipeline from Reader
        /// </summary>
        public static StreamingPipeline<T> CreatePipeline<T>(
            IStreamingReader reader,
            IRowMapper<T> mapper) where T : new()
        {
            return new StreamingPipeline<T>(reader, mapper);
        }
    }
}
