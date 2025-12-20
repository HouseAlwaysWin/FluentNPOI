using FluentNPOI.Charts;
using Xunit;
using System.Collections.Generic;
using System.Linq;

namespace FluentNPOIUnitTest
{
    public class ChartBuilderTests
    {
        private class TestData
        {
            public string Category { get; set; } = "";
            public double Value { get; set; }
        }

        private List<TestData> GetTestData() => new List<TestData>
        {
            new TestData { Category = "A", Value = 10 },
            new TestData { Category = "B", Value = 25 },
            new TestData { Category = "C", Value = 15 },
            new TestData { Category = "D", Value = 30 },
            new TestData { Category = "E", Value = 20 }
        };

        [Fact]
        public void BarChart_ShouldGeneratePngBytes()
        {
            // Arrange
            var data = GetTestData();

            // Act
            var bytes = ChartBuilder.Bar(data)
                .X(d => d.Category)
                .Y(d => d.Value)
                .WithTitle("Test Bar Chart")
                .ToPng(400, 300);

            // Assert
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
            // PNG magic bytes: 89 50 4E 47
            Assert.Equal(0x89, bytes[0]);
            Assert.Equal(0x50, bytes[1]);
            Assert.Equal(0x4E, bytes[2]);
            Assert.Equal(0x47, bytes[3]);
        }

        [Fact]
        public void LineChart_ShouldGeneratePngBytes()
        {
            // Arrange
            var data = GetTestData();

            // Act
            var bytes = ChartBuilder.Line(data)
                .Y(d => d.Value)
                .WithTitle("Test Line Chart")
                .WithXLabel("Index")
                .WithYLabel("Value")
                .ToPng(400, 300);

            // Assert
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void PieChart_ShouldGeneratePngBytes()
        {
            // Arrange
            var data = GetTestData();

            // Act
            var bytes = ChartBuilder.Pie(data)
                .X(d => d.Category)
                .Y(d => d.Value)
                .WithTitle("Test Pie Chart")
                .ToPng(400, 400);

            // Assert
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void ScatterChart_ShouldGeneratePngBytes()
        {
            // Arrange
            var data = GetTestData();

            // Act
            var bytes = ChartBuilder.Scatter(data)
                .Y(d => d.Value)
                .WithTitle("Test Scatter Chart")
                .ToPng(400, 300);

            // Assert
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void Configure_ShouldAllowScottPlotCustomization()
        {
            // Arrange
            var data = GetTestData();
            bool configureWasCalled = false;

            // Act
            var bytes = ChartBuilder.Bar(data)
                .X(d => d.Category)
                .Y(d => d.Value)
                .Configure(plot =>
                {
                    configureWasCalled = true;
                    plot.FigureBackground.Color = ScottPlot.Colors.White;
                })
                .ToPng(400, 300);

            // Assert
            Assert.True(configureWasCalled);
            Assert.NotNull(bytes);
        }

        [Fact]
        public void ChartBuilder_WithEmptyData_ShouldNotThrow()
        {
            // Arrange
            var emptyData = new List<TestData>();

            // Act & Assert - Should not throw
            var bytes = ChartBuilder.Bar(emptyData)
                .X(d => d.Category)
                .Y(d => d.Value)
                .ToPng(400, 300);

            Assert.NotNull(bytes);
        }
    }
}
