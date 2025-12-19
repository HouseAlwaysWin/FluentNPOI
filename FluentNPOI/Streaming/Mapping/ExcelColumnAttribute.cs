using System;

namespace FluentNPOI.Streaming.Mapping
{
    /// <summary>
    /// Excel column mapping attribute (optional)
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Column index (0-based)
        /// </summary>
        public int Index { get; set; } = -1;

        /// <summary>
        /// Column name (Header mapping)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Title (used when writing)
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Format
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Create with index
        /// </summary>
        public ExcelColumnAttribute(int index)
        {
            Index = index;
        }

        /// <summary>
        /// Create with name
        /// </summary>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelColumnAttribute()
        {
        }
    }
}
