# FluentNPOI

[![CI](https://github.com/HouseAlwaysWin/FluentNPOI/workflows/CI/badge.svg)](https://github.com/HouseAlwaysWin/FluentNPOI/actions/workflows/ci.yml)
[![.NET Standard 2.0](https://img.shields.io/badge/.NET%20Standard-2.0-blue.svg)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**FluentNPOI** 是基於 [NPOI](https://github.com/dotnetcore/NPOI) 的流暢（Fluent）風格 Excel 操作庫，提供更直觀、更易用的 API 來讀寫 Excel 文件。

[English](#english) | [繁體中文](#繁體中文)

---

## 繁體中文

### 🚀 特性

- ✅ **流暢 API** - 支援鏈式調用，代碼更簡潔易讀
- ✅ **強型別映射** - 透過 `FluentMapping` 進行強型別資料綁定與樣式設定
- ✅ **模組化套件** - 按需安裝：核心、PDF、串流、圖表
- ✅ **直觀樣式** - 支援在 Mapping 中直接設定樣式，或使用 FluentCell API 進行細粒度控制
- ✅ **樣式管理** - 智能樣式緩存機制，自動處理重複樣式
- ✅ **完整讀寫** - 支援讀寫 Excel、圖片插入、公式設定、合併儲存格
- ✅ **HTML/PDF 匯出** - 將 Excel 轉換為 HTML 或 PDF
- ✅ **圖表產生** - 使用 ScottPlot 產生圖表並嵌入 Excel

### 📦 安裝

#### 核心套件

```bash
dotnet add package FluentNPOI
```

#### 可選模組

| 套件 | 用途 | 安裝 |
|------|------|------|
| `FluentNPOI.Pdf` | PDF 匯出 (QuestPDF) | `dotnet add package FluentNPOI.Pdf` |
| `FluentNPOI.Streaming` | 大檔案串流讀寫 | `dotnet add package FluentNPOI.Streaming` |
| `FluentNPOI.Charts` | 圖表產生 (ScottPlot) | `dotnet add package FluentNPOI.Charts` |
| `FluentNPOI.All` | 完整功能 (包含所有模組) | `dotnet add package FluentNPOI.All` |

### 🎯 快速開始

#### 1. 基本讀寫

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var fluent = new FluentWorkbook(workbook);

fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.A, 1)
      .SetValue("Hello World!")
      .SetBackgroundColor(IndexedColors.Yellow)
      .SetFont(isBold: true, fontSize: 14);

fluent.SaveToPath("output.xlsx");
```

#### 2. 強型別表格綁定 (推薦)

```csharp
var mapping = new FluentMapping<Student>();

mapping.Map(x => x.Name)
    .ToColumn(ExcelCol.A)
    .WithTitle("姓名")
    .WithBackgroundColor(IndexedColors.LightCornflowerBlue);

mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithTitle("分數")
    .WithNumberFormat("0.0");

fluent.UseSheet("Report")
      .SetTable(data, mapping)
      .BuildRows()
      .SetAutoFilter()
      .FreezeTitleRow();
```

#### 3. 串流處理大檔案

```csharp
using FluentNPOI.Streaming;

StreamingBuilder<DataModel>.FromFile("large_input.xlsx")
    .Transform(x => x.Value *= 2)
    .WithMapping(mapping)
    .SaveAs("output.xlsx");
```

#### 4. 圖表產生

```csharp
using FluentNPOI.Charts;

// 整合串鍊 API
fluent.UseSheet("Charts")
    .SetCellPosition(ExcelCol.A, 1)
    .AddBarChart(data, chart => {
        chart.X(d => d.Category)
             .Y(d => d.Value)
             .WithTitle("Sales Report");
    }, width: 400, height: 300);

// 或手動產生
var chartBytes = ChartBuilder.Bar(data)
    .X(d => d.Category)
    .Y(d => d.Value)
    .Configure(plot => {
        // 完整存取 ScottPlot API
        plot.FigureBackground.Color = ScottPlot.Colors.White;
    })
    .ToPng(400, 300);
```

#### 5. PDF 匯出

```csharp
using FluentNPOI.Pdf;

PdfConverter.ConvertSheetToPdf(fluent.UseSheet("Report"), "report.pdf");
```

### 📖 API 概覽

| 用途 | 主要方法 |
|------|----------|
| **Mapping** | `Map`, `ToColumn`, `WithTitle`, `WithNumberFormat`, `WithBackgroundColor` |
| **Cell** | `SetValue`, `SetFormula`, `SetBackgroundColor`, `SetBorder`, `SetFont` |
| **Table** | `SetTable`, `BuildRows`, `SetAutoFilter`, `FreezeTitleRow`, `AutoSizeColumns` |
| **Streaming** | `StreamingBuilder.FromFile`, `Transform`, `SaveAs` |
| **Charts** | `AddBarChart`, `AddLineChart`, `AddPieChart`, `ChartBuilder` |

---

## English

### 🚀 Features

- ✅ **Fluent API** - Chained method calls for simpler, readable code
- ✅ **Strongly Typed Mapping** - Use `FluentMapping` for type-safe data binding
- ✅ **Modular Packages** - Install only what you need: Core, PDF, Streaming, Charts
- ✅ **Direct Styling** - Configure styles directly within Mapping or FluentCell API
- ✅ **Style Management** - Smart caching to handle duplicate styles
- ✅ **Comprehensive I/O** - Read/Write, Images, Formulas, Merging
- ✅ **HTML/PDF Export** - Convert Excel to HTML or PDF
- ✅ **Chart Generation** - Generate charts using ScottPlot and embed in Excel

### 📦 Installation

#### Core Package

```bash
dotnet add package FluentNPOI
```

#### Optional Modules

| Package | Purpose | Install |
|---------|---------|---------|
| `FluentNPOI.Pdf` | PDF Export (QuestPDF) | `dotnet add package FluentNPOI.Pdf` |
| `FluentNPOI.Streaming` | Large File Streaming | `dotnet add package FluentNPOI.Streaming` |
| `FluentNPOI.Charts` | Chart Generation (ScottPlot) | `dotnet add package FluentNPOI.Charts` |
| `FluentNPOI.All` | Full Features (All modules) | `dotnet add package FluentNPOI.All` |

### 🎯 Quick Start

#### 1. Basic Write

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var fluent = new FluentWorkbook(workbook);

fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.A, 1)
      .SetValue("Hello World!")
      .SetBackgroundColor(IndexedColors.Yellow)
      .SetFont(isBold: true, fontSize: 14);

fluent.SaveToPath("output.xlsx");
```

#### 2. Table Binding with FluentMapping (Recommended)

```csharp
var mapping = new FluentMapping<Student>();

mapping.Map(x => x.Name)
    .ToColumn(ExcelCol.A)
    .WithTitle("Name")
    .WithBackgroundColor(IndexedColors.LightCornflowerBlue);

mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithTitle("Score")
    .WithNumberFormat("0.0");

fluent.UseSheet("Report")
      .SetTable(data, mapping)
      .BuildRows()
      .SetAutoFilter()
      .FreezeTitleRow();
```

#### 3. Streaming for Large Files

```csharp
using FluentNPOI.Streaming;

StreamingBuilder<DataModel>.FromFile("large_input.xlsx")
    .Transform(x => x.Value *= 2)
    .WithMapping(mapping)
    .SaveAs("output.xlsx");
```

#### 4. Chart Generation

```csharp
using FluentNPOI.Charts;

// Integrated chaining API
fluent.UseSheet("Charts")
    .SetCellPosition(ExcelCol.A, 1)
    .AddBarChart(data, chart => {
        chart.X(d => d.Category)
             .Y(d => d.Value)
             .WithTitle("Sales Report");
    }, width: 400, height: 300);

// Or generate manually
var chartBytes = ChartBuilder.Bar(data)
    .X(d => d.Category)
    .Y(d => d.Value)
    .Configure(plot => {
        // Full access to ScottPlot API
        plot.FigureBackground.Color = ScottPlot.Colors.White;
    })
    .ToPng(400, 300);
```

#### 5. PDF Export

```csharp
using FluentNPOI.Pdf;

PdfConverter.ConvertSheetToPdf(fluent.UseSheet("Report"), "report.pdf");
```

### 📖 API Overview

| Area | Key Methods |
|------|-------------|
| **Mapping** | `Map`, `ToColumn`, `WithTitle`, `WithNumberFormat`, `WithBackgroundColor` |
| **Cell** | `SetValue`, `SetFormula`, `SetBackgroundColor`, `SetBorder`, `SetFont` |
| **Table** | `SetTable`, `BuildRows`, `SetAutoFilter`, `FreezeTitleRow`, `AutoSizeColumns` |
| **Streaming** | `StreamingBuilder.FromFile`, `Transform`, `SaveAs` |
| **Charts** | `AddBarChart`, `AddLineChart`, `AddPieChart`, `ChartBuilder` |

---

### 🤝 Contribution

Issues and Pull Requests are welcome!

### 📄 License

MIT License - See [LICENSE](LICENSE) file.
