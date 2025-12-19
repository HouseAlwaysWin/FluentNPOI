using NPOI.SS.UserModel;
using System;

namespace FluentNPOI.Models
{
    /// <summary>
    /// Cell style configuration
    /// </summary>
    public class CellStyleConfig
    {
        /// <summary>
        /// Style cache key (same key reuses style)
        /// If null or empty string, style will not be cached
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Style setter (executed only when style not in cache)
        /// </summary>
        public Action<ICellStyle> StyleSetter { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="key">Style cache key</param>
        /// <param name="styleSetter">Style setter</param>
        public CellStyleConfig(string key, Action<ICellStyle> styleSetter)
        {
            Key = key;
            StyleSetter = styleSetter;
        }

        /// <summary>
        /// Deconstruct method (supports tuple syntax)
        /// </summary>
        public void Deconstruct(out string key, out Action<ICellStyle> styleSetter)
        {
            key = Key;
            styleSetter = StyleSetter;
        }
    }
}

