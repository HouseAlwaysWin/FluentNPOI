using ScottPlot;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOI.Charts
{
    /// <summary>
    /// Builder for creating charts using ScottPlot and exporting as images
    /// </summary>
    /// <typeparam name="T">Data item type</typeparam>
    public class ChartBuilder<T>
    {
        private readonly IEnumerable<T> _data;
        private readonly ChartType _chartType;
        private Func<T, string> _xSelector;
        private Func<T, double> _ySelector;
        private string _title;
        private string _xLabel;
        private string _yLabel;
        private Action<Plot> _configurePlot;

        private ChartBuilder(IEnumerable<T> data, ChartType chartType)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
            _chartType = chartType;
        }

        /// <summary>
        /// Create a bar chart from data
        /// </summary>
        public static ChartBuilder<T> Bar(IEnumerable<T> data) => new ChartBuilder<T>(data, ChartType.Bar);

        /// <summary>
        /// Create a line chart from data
        /// </summary>
        public static ChartBuilder<T> Line(IEnumerable<T> data) => new ChartBuilder<T>(data, ChartType.Line);

        /// <summary>
        /// Create a scatter chart from data
        /// </summary>
        public static ChartBuilder<T> Scatter(IEnumerable<T> data) => new ChartBuilder<T>(data, ChartType.Scatter);

        /// <summary>
        /// Create a pie chart from data
        /// </summary>
        public static ChartBuilder<T> Pie(IEnumerable<T> data) => new ChartBuilder<T>(data, ChartType.Pie);

        /// <summary>
        /// Set X-axis values (labels for bar/pie, numeric for line/scatter)
        /// </summary>
        public ChartBuilder<T> X(Func<T, string> selector)
        {
            _xSelector = selector;
            return this;
        }

        /// <summary>
        /// Set Y-axis values
        /// </summary>
        public ChartBuilder<T> Y(Func<T, double> selector)
        {
            _ySelector = selector;
            return this;
        }

        /// <summary>
        /// Set chart title
        /// </summary>
        public ChartBuilder<T> WithTitle(string title)
        {
            _title = title;
            return this;
        }

        /// <summary>
        /// Set X-axis label
        /// </summary>
        public ChartBuilder<T> WithXLabel(string label)
        {
            _xLabel = label;
            return this;
        }

        /// <summary>
        /// Set Y-axis label
        /// </summary>
        public ChartBuilder<T> WithYLabel(string label)
        {
            _yLabel = label;
            return this;
        }

        /// <summary>
        /// Configure the ScottPlot Plot object directly for advanced customization
        /// </summary>
        /// <param name="configure">Action to configure the Plot object</param>
        /// <example>
        /// ChartBuilder.Bar(data)
        ///     .X(d => d.Name)
        ///     .Y(d => d.Value)
        ///     .Configure(plot => {
        ///         plot.Style.Background(Color.FromHex("#1a1a1a"));
        ///         plot.Axes.Left.Label.ForeColor = Colors.White;
        ///     })
        ///     .ToPng(400, 300);
        /// </example>
        public ChartBuilder<T> Configure(Action<Plot> configure)
        {
            _configurePlot = configure;
            return this;
        }

        /// <summary>
        /// Generate chart as PNG byte array
        /// </summary>
        /// <param name="width">Image width in pixels</param>
        /// <param name="height">Image height in pixels</param>
        /// <returns>PNG image bytes</returns>
        public byte[] ToPng(int width = 400, int height = 300)
        {
            var plot = new Plot();

            var dataList = _data.ToList();
            var labels = _xSelector != null ? dataList.Select(_xSelector).ToArray() : null;
            var values = _ySelector != null ? dataList.Select(_ySelector).ToArray() : new double[0];

            switch (_chartType)
            {
                case ChartType.Bar:
                    BuildBarChart(plot, labels, values);
                    break;
                case ChartType.Line:
                    BuildLineChart(plot, values);
                    break;
                case ChartType.Scatter:
                    BuildScatterChart(plot, values);
                    break;
                case ChartType.Pie:
                    BuildPieChart(plot, labels, values);
                    break;
            }

            // Automatic font detection for CJK (Chinese/Japanese/Korean) support
            plot.Font.Automatic();

            if (!string.IsNullOrEmpty(_title))
                plot.Title(_title);

            if (!string.IsNullOrEmpty(_xLabel))
                plot.XLabel(_xLabel);

            if (!string.IsNullOrEmpty(_yLabel))
                plot.YLabel(_yLabel);

            // Apply custom configuration
            _configurePlot?.Invoke(plot);

            return plot.GetImageBytes(width, height, ImageFormat.Png);
        }

        private void BuildBarChart(Plot plot, string[] labels, double[] values)
        {
            var positions = Enumerable.Range(0, values.Length).Select(i => (double)i).ToArray();
            var bars = plot.Add.Bars(positions, values);

            if (labels != null && labels.Length > 0)
            {
                var ticks = positions.Select((p, i) => new Tick(p, labels[i])).ToArray();
                plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.NumericManual(ticks);
                plot.Axes.Bottom.MajorTickStyle.Length = 0;
            }
        }

        private void BuildLineChart(Plot plot, double[] values)
        {
            var xs = Enumerable.Range(0, values.Length).Select(i => (double)i).ToArray();
            plot.Add.Scatter(xs, values);
        }

        private void BuildScatterChart(Plot plot, double[] values)
        {
            var xs = Enumerable.Range(0, values.Length).Select(i => (double)i).ToArray();
            var scatter = plot.Add.Scatter(xs, values);
            scatter.LineWidth = 0; // Scatter only shows markers
        }

        private void BuildPieChart(Plot plot, string[] labels, double[] values)
        {
            var slices = new List<PieSlice>();
            for (int i = 0; i < values.Length; i++)
            {
                var label = labels != null && i < labels.Length ? labels[i] : $"Item {i + 1}";
                slices.Add(new PieSlice { Value = values[i], Label = label });
            }
            plot.Add.Pie(slices);
        }
    }

    /// <summary>
    /// Static factory for creating charts
    /// </summary>
    public static class ChartBuilder
    {
        /// <summary>Create a bar chart</summary>
        public static ChartBuilder<T> Bar<T>(IEnumerable<T> data) => ChartBuilder<T>.Bar(data);

        /// <summary>Create a line chart</summary>
        public static ChartBuilder<T> Line<T>(IEnumerable<T> data) => ChartBuilder<T>.Line(data);

        /// <summary>Create a scatter chart</summary>
        public static ChartBuilder<T> Scatter<T>(IEnumerable<T> data) => ChartBuilder<T>.Scatter(data);

        /// <summary>Create a pie chart</summary>
        public static ChartBuilder<T> Pie<T>(IEnumerable<T> data) => ChartBuilder<T>.Pie(data);
    }
}
