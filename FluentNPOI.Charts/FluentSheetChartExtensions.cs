using FluentNPOI.Stages;
using FluentNPOI.Models;
using System;
using System.Collections.Generic;

namespace FluentNPOI.Charts
{
    /// <summary>
    /// Extension methods for FluentCell to add charts
    /// </summary>
    public static class FluentCellChartExtensions
    {
        /// <summary>
        /// Add a bar chart at current cell position
        /// </summary>
        /// <typeparam name="T">Data type</typeparam>
        /// <param name="cell">FluentCell instance</param>
        /// <param name="data">Chart data</param>
        /// <param name="configure">Chart configuration action</param>
        /// <param name="width">Chart width in pixels</param>
        /// <param name="height">Chart height in pixels</param>
        /// <returns>FluentCell for chaining</returns>
        public static FluentCell AddBarChart<T>(
            this FluentCell cell,
            IEnumerable<T> data,
            Action<ChartBuilder<T>> configure,
            int width = 400,
            int height = 300)
        {
            var builder = ChartBuilder<T>.Bar(data);
            configure?.Invoke(builder);
            var bytes = builder.ToPng(width, height);
            return cell.SetPictureOnCell(bytes, width, height);
        }

        /// <summary>
        /// Add a line chart at current cell position
        /// </summary>
        public static FluentCell AddLineChart<T>(
            this FluentCell cell,
            IEnumerable<T> data,
            Action<ChartBuilder<T>> configure,
            int width = 400,
            int height = 300)
        {
            var builder = ChartBuilder<T>.Line(data);
            configure?.Invoke(builder);
            var bytes = builder.ToPng(width, height);
            return cell.SetPictureOnCell(bytes, width, height);
        }

        /// <summary>
        /// Add a scatter chart at current cell position
        /// </summary>
        public static FluentCell AddScatterChart<T>(
            this FluentCell cell,
            IEnumerable<T> data,
            Action<ChartBuilder<T>> configure,
            int width = 400,
            int height = 300)
        {
            var builder = ChartBuilder<T>.Scatter(data);
            configure?.Invoke(builder);
            var bytes = builder.ToPng(width, height);
            return cell.SetPictureOnCell(bytes, width, height);
        }

        /// <summary>
        /// Add a pie chart at current cell position
        /// </summary>
        public static FluentCell AddPieChart<T>(
            this FluentCell cell,
            IEnumerable<T> data,
            Action<ChartBuilder<T>> configure,
            int width = 400,
            int height = 300)
        {
            var builder = ChartBuilder<T>.Pie(data);
            configure?.Invoke(builder);
            var bytes = builder.ToPng(width, height);
            return cell.SetPictureOnCell(bytes, width, height);
        }
    }
}
