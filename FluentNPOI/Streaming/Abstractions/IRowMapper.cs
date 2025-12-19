namespace FluentNPOI.Streaming.Abstractions
{
    /// <summary>
    /// Interface for mapping row data to DTO
    /// </summary>
    /// <typeparam name="T">Target DTO type</typeparam>
    public interface IRowMapper<T>
    {
        /// <summary>
        /// Convert streaming row to DTO
        /// </summary>
        /// <param name="row">Streaming row data</param>
        /// <returns>Converted DTO</returns>
        T Map(IStreamingRow row);
    }
}
