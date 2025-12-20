# FluentNPOI.Charts

Chart generation extension for FluentNPOI using ScottPlot.

## Installation

```bash
dotnet add package FluentNPOI.Charts
```

## Usage

### Integrated Chaining (Recommended)

```csharp
using FluentNPOI.Charts;

fluent.UseSheet("Charts")
    .SetCellPosition(ExcelCol.A, 1)
    .AddBarChart(data, chart => {
        chart.X(d => d.Month)
             .Y(d => d.Revenue)
             .WithTitle("Monthly Revenue");
    }, width: 400, height: 300)
    .SetCellPosition(ExcelCol.G, 1)
    .AddLineChart(data, chart => {
        chart.Y(d => d.Growth)
             .WithTitle("Growth Trend");
    });
```

### Manual Generation

```csharp
var chartBytes = ChartBuilder.Bar(data)
    .X(d => d.Month)
    .Y(d => d.Revenue)
    .ToPng(400, 300);

sheet.SetCellPosition(ExcelCol.A, 1)
     .SetPictureOnCell(chartBytes, 400, 300);
```

## Supported Chart Types

| Type | Method |
|------|--------|
| Bar | `AddBarChart()` |
| Line | `AddLineChart()` |
| Scatter | `AddScatterChart()` |
| Pie | `AddPieChart()` |
