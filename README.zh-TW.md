# FluentNPOI

[![CI](https://github.com/HouseAlwaysWin/FluentNPOI/workflows/CI/badge.svg)](https://github.com/HouseAlwaysWin/FluentNPOI/actions/workflows/ci.yml)
[![.NET Standard 2.0](https://img.shields.io/badge/.NET%20Standard-2.0-blue.svg)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**FluentNPOI** 是基於 [NPOI](https://github.com/dotnetcore/NPOI) 的流暢（Fluent）風格 Excel 操作庫，提供更直觀、更易用的 API 來讀寫 Excel 文件。

[ English ](README.md)

---

## 🚀 特性

- ✅ **流暢 API** - 支援鏈式調用，代碼更簡潔易讀
- ✅ **強型別映射** - 透過 `FluentMapping` 進行強型別資料綁定與樣式設定
- ✅ **模組化套件** - 按需安裝：核心、PDF、串流、圖表
- ✅ **直觀樣式** - 支援在 Mapping 中直接設定樣式，或使用 FluentCell API 進行細粒度控制
- ✅ **樣式管理** - 智能樣式緩存機制，自動處理重複樣式
- ✅ **完整讀寫** - 支援讀寫 Excel、圖片插入、公式設定、合併儲存格
- ✅ **HTML/PDF 匯出** - 將 Excel 轉換為 HTML 或 PDF
- ✅ **圖表產生** - 使用 ScottPlot 產生圖表並嵌入 Excel
- ✅ **即時預覽 (Hot Reload)** - 支援 `dotnet watch` 與 LibreOffice 即時預覽變更 (需安裝 LibreOffice)

## 📦 安裝

### 核心套件

```bash
dotnet add package FluentNPOI
```

### 可選模組

| 套件 | 用途 | 安裝 |
|------|------|------|
| `FluentNPOI.Pdf` | PDF 匯出 (QuestPDF) | `dotnet add package FluentNPOI.Pdf` |
| `FluentNPOI.Streaming` | 大檔案串流讀寫 | `dotnet add package FluentNPOI.Streaming` |
| `FluentNPOI.Charts` | 圖表產生 (ScottPlot) | `dotnet add package FluentNPOI.Charts` |
| `FluentNPOI.HotReload` | 即時預覽 (開發用) | `dotnet add package FluentNPOI.HotReload` |
| `FluentNPOI.All` | 完整功能 (包含所有模組) | `dotnet add package FluentNPOI.All` |

## 🎯 快速開始

### 1. 基本讀寫

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

### 2. 強型別表格綁定 (推薦)

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

### 3. 串流處理大檔案

```csharp
using FluentNPOI.Streaming;

StreamingBuilder<DataModel>.FromFile("large_input.xlsx")
    .Transform(x => x.Value *= 2)
    .WithMapping(mapping)
    .SaveAs("output.xlsx");
```

### 4. 圖表產生

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

### 5. PDF 匯出

```csharp
using FluentNPOI.Pdf;

PdfConverter.ConvertSheetToPdf(fluent.UseSheet("Report"), "report.pdf");
```

### 6. 即時預覽 (Hot Reload)

確保已安裝 `FluentNPOI.HotReload` 與 LibreOffice。

#### 程式碼實作

使用 `FluentLivePreview.Run` 包裝您的產生邏輯：

```csharp
using FluentNPOI.HotReload;

// ... 在 Main 方法或設定中
FluentLivePreview.Run("output/report.xlsx", fluent =>
{
    // 在此撰寫 FluentNPOI 程式碼
    fluent.UseSheet("Sheet1")
          .SetCellPosition(ExcelCol.A, 1)
          .SetValue("即時更新！")
          .SetBackgroundColor(IndexedColors.LightGreen);
});
```

#### 使用 dotnet watch 執行

```bash
# 在 Console 專案目錄下執行
dotnet watch run
```

修改代碼後儲存，LibreOffice 將會自動重新載入並顯示最新結果。

## 📖 API 概覽

| 用途 | 主要方法 |
|------|----------|
| **Mapping** | `Map`, `ToColumn`, `WithTitle`, `WithNumberFormat`, `WithBackgroundColor` |
| **Cell** | `SetValue`, `SetFormula`, `SetBackgroundColor`, `SetBorder`, `SetFont` |
| **Table** | `SetTable`, `BuildRows`, `SetAutoFilter`, `FreezeTitleRow`, `AutoSizeColumns` |
| **Streaming** | `StreamingBuilder.FromFile`, `Transform`, `SaveAs` |
| **Charts** | `AddBarChart`, `AddLineChart`, `AddPieChart`, `ChartBuilder` |
| **HotReload** | `FluentLivePreview.Run` |

---

### 🤝 貢獻

歡迎提交 Issues 和 Pull Requests！

### 📄 授權

MIT License - 詳見 [LICENSE](LICENSE) 文件。
