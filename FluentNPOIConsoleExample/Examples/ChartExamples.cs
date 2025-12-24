using System;
using System.Linq;
using System.Collections.Generic;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;
using FluentNPOI.Charts;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Chart Examples - ScottPlot integration
    /// </summary>
    internal partial class Program
    {
        #region Chart Examples

        /// <summary>
        /// Example 14: Generate and embed charts using ScottPlot
        /// </summary>
        public static void CreateChartExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("Creating ChartExample...");

            var sheet = fluent.UseSheet("ChartExample", true);
            sheet.SetColumnWidth(ExcelCol.A, ExcelCol.L, 15);

            // Title
            sheet.SetCellPosition(ExcelCol.A, 1)
                 .SetValue("Chart Examples - ScottPlot Integration")
                 .SetCellStyle("HeaderBlue");
            sheet.SetExcelCellMerge(ExcelCol.A, ExcelCol.F, 1);

            // 1. Bar Chart - Score distribution
            var barChartBytes = ChartBuilder.Bar(testData)
                .X(d => d.Name.Length > 10 ? d.Name.Substring(0, 10) + "..." : d.Name)
                .Y(d => d.Score)
                .WithTitle("Scores by Name")
                .WithYLabel("Score")
                .ToPng(400, 300);

            sheet.SetCellPosition(ExcelCol.A, 3)
                 .SetValue("Bar Chart: Scores")
                 .SetCellStyle("HighlightYellow");

            sheet.SetCellPosition(ExcelCol.A, 4)
                 .SetPictureOnCell(barChartBytes, 400, 300);

            // 2. Line Chart - Score trend
            var lineChartBytes = ChartBuilder.Line(testData)
                .Y(d => d.Score)
                .WithTitle("Score Trend")
                .WithXLabel("Index")
                .WithYLabel("Score")
                .ToPng(400, 300);

            sheet.SetCellPosition(ExcelCol.G, 3)
                 .SetValue("Line Chart: Score Trend")
                 .SetCellStyle("HighlightYellow");

            sheet.SetCellPosition(ExcelCol.G, 4)
                 .SetPictureOnCell(lineChartBytes, 400, 300);

            // 3. Pie Chart - Score Distribution
            var scoreRanges = new[]
            {
                new { Label = "90+ 優秀", Value = (double)testData.Count(d => d.Score >= 90), Color = "#4CAF50" },
                new { Label = "80-89 良好", Value = (double)testData.Count(d => d.Score >= 80 && d.Score < 90), Color = "#8BC34A" },
                new { Label = "70-79 普通", Value = (double)testData.Count(d => d.Score >= 70 && d.Score < 80), Color = "#FFC107" },
                new { Label = "60-69 及格", Value = (double)testData.Count(d => d.Score >= 60 && d.Score < 70), Color = "#FF9800" },
                new { Label = "<60 不及格", Value = (double)testData.Count(d => d.Score < 60), Color = "#F44336" }
            };

            sheet.SetCellPosition(ExcelCol.A, 20)
                 .SetValue("🥧 Pie Chart: Score Distribution (Legend Style)")
                 .SetCellStyle("HighlightYellow");

            var pieChartBytes = ChartBuilder.Pie(scoreRanges)
                .X(d => d.Label)
                .Y(d => d.Value)
                .WithTitle("Score Distribution")
                .Configure(plot =>
                {
                    plot.Axes.Frameless();
                    plot.HideGrid();

                    var pie = plot.GetPlottables().OfType<ScottPlot.Plottables.Pie>().FirstOrDefault();
                    if (pie != null)
                    {
                        double total = scoreRanges.Sum(r => r.Value);
                        for (int i = 0; i < pie.Slices.Count && i < scoreRanges.Length; i++)
                        {
                            pie.Slices[i].FillColor = ScottPlot.Color.FromHex(scoreRanges[i].Color);
                            pie.Slices[i].LabelFontColor = ScottPlot.Colors.Black;
                            pie.Slices[i].LabelFontSize = 11;

                            double pct = scoreRanges[i].Value / total * 100;
                            pie.Slices[i].Label = $"{scoreRanges[i].Label}\n({pct:F0}%)";
                        }
                        pie.ExplodeFraction = 0.03;
                    }

                    plot.FigureBackground.Color = ScottPlot.Colors.White;
                    plot.DataBackground.Color = ScottPlot.Colors.White;
                    plot.Title("Score Distribution", size: 18);
                })
                .ToPng(500, 450);

            sheet.SetCellPosition(ExcelCol.A, 21)
                 .SetPictureOnCell(pieChartBytes, 500, 450);

            sheet.SetCellPosition(ExcelCol.A, 40)
                 .SetValue("💡 提示：超過 10 個項目建議使用 Bar Chart");

            // 4. Custom Styled Chart
            sheet.SetCellPosition(ExcelCol.G, 20)
                 .SetValue("Custom Styled: ScottPlot Configure")
                 .SetCellStyle("HighlightYellow");

            var customChartBytes = ChartBuilder.Bar(testData.Take(5))
                .X(d => d.Name.Length > 8 ? d.Name.Substring(0, 8) : d.Name)
                .Y(d => d.Score)
                .WithTitle("Custom Styled Bar Chart")
                .Configure(plot =>
                {
                    plot.FigureBackground.Color = ScottPlot.Color.FromHex("#2d2d30");
                    plot.DataBackground.Color = ScottPlot.Color.FromHex("#1e1e1e");
                    plot.Axes.Color(ScottPlot.Color.FromHex("#d4d4d4"));
                    plot.Legend.IsVisible = false;
                })
                .ToPng(400, 300);

            sheet.SetCellPosition(ExcelCol.G, 21)
                 .SetPictureOnCell(customChartBytes, 400, 300);

            // 5. Large Data Bar Chart
            sheet.SetCellPosition(ExcelCol.A, 42)
                 .SetValue("📊 Bar Chart: 30+ Items (Recommended for large data)")
                 .SetCellStyle("HighlightYellow");

            var largeData = Enumerable.Range(1, 35).Select(i => new
            {
                Category = $"Category {i:D2}",
                Value = 50 + Math.Sin(i * 0.3) * 40 + (i % 5) * 10
            }).ToArray();

            var largeBarChartBytes = ChartBuilder.Bar(largeData)
                .X(d => d.Category)
                .Y(d => d.Value)
                .WithTitle("35 Categories - Bar Chart")
                .WithYLabel("Value")
                .Configure(plot =>
                {
                    plot.FigureBackground.Color = ScottPlot.Colors.White;
                    plot.DataBackground.Color = ScottPlot.Color.FromHex("#f5f5f5");

                    var bars = plot.GetPlottables().OfType<ScottPlot.Plottables.BarPlot>().FirstOrDefault();
                    if (bars != null)
                    {
                        foreach (var bar in bars.Bars)
                        {
                            var intensity = Math.Min(1.0, bar.Value / 100.0);
                            bar.FillColor = ScottPlot.Color.FromHex(intensity > 0.7 ? "#4CAF50" :
                                                                     intensity > 0.5 ? "#8BC34A" :
                                                                     intensity > 0.3 ? "#FFC107" : "#FF9800");
                        }
                    }

                    plot.Axes.Bottom.TickLabelStyle.Rotation = 45;
                    plot.Axes.Bottom.TickLabelStyle.Alignment = ScottPlot.Alignment.MiddleLeft;
                    plot.Grid.MajorLineColor = ScottPlot.Color.FromHex("#e0e0e0");
                })
                .ToPng(800, 500);

            sheet.SetCellPosition(ExcelCol.A, 43)
                 .SetPictureOnCell(largeBarChartBytes, 800, 500);

            Console.WriteLine("  ✓ ChartExample Created");
        }

        #endregion
    }
}
