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
- ✅ **直觀樣式** - 支援在 Mapping 中直接設定樣式，或使用 FluentCell API 進行細粒度控制
- ✅ **樣式管理** - 智能樣式緩存機制，自動處理重複樣式，避免 Excel 樣式上限問題
- ✅ **完整讀寫** - 支援讀寫 Excel、圖片插入、公式設定、合併儲存格
- ✅ **工作簿管理** - 支援工作表複製、刪除、重新命名、調整行高列寬

### 📦 安裝

```bash
# 使用 NuGet Package Manager
Install-Package FluentNPOI

# 使用 .NET CLI
dotnet add package FluentNPOI
```

### 🎯 快速開始

#### 1. 基本讀寫

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var fluent = new FluentWorkbook(workbook);

fluent.UseSheet("Sheet1")
      // 定位並鏈式操作
      .SetCellPosition(0, 1) // Row 1, Col A (0-based)
      .SetValue("Hello World!")
      .SetBackgroundColor(IndexedColors.Yellow) // 直接設定樣式
      .SetFont(isBold: true, fontSize: 14)
      .SetAlignment(HorizontalAlignment.Center);

fluent.SaveToPath("output.xlsx");
```

#### 2. 強型別表格綁定與樣式 (推薦)

使用 `FluentMapping` 可以同時定義資料與外觀，這是處理列表資料最推薦的方式。

```csharp
var data = new List<Student>
{
    new Student { Name = "Alice", Score = 95.5, Date = DateTime.Now },
    new Student { Name = "Bob", Score = 80.0, Date = DateTime.Now }
};

// 定義 Mapping 與樣式
var mapping = new FluentMapping<Student>();

mapping.Map(x => x.Name)
    .ToColumn(ExcelCol.A)
    .WithTitle("姓名")
    .WithAlignment(HorizontalAlignment.Center)
    .WithBackgroundColor(IndexedColors.LightCornflowerBlue);

mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithTitle("分數")
    .WithNumberFormat("0.0") // 設定數值格式
    .WithFont(isBold: true);

mapping.Map(x => x.Date)
    .ToColumn(ExcelCol.C)
    .WithTitle("日期")
    .WithNumberFormat("yyyy-mm-dd");

// 寫入並應用功能
fluent.UseSheet("Report")
      .SetTable(data, mapping)
      .BuildRows()
      .SetAutoFilter() // 自動篩選
      .FreezeTitleRow() // 凍結標題行
      .AutoSizeColumns(); // 自動調整欄寬
```

### 📚 主要功能

#### 1. 單元格操作 (FluentCell)

FluentCell 提供了豐富的鏈式方法來操作單元格：

```csharp
fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.C, 5)
      .SetValue(12345.678)
      .SetNumberFormat("#,##0.00")         // 數值格式
      .SetBackgroundColor(IndexedColors.Red) // 背景色
      .SetFont(fontName: "Arial", isBold: true) // 字體
      .SetBorder(BorderStyle.Thin)         // 邊框
      .SetAlignment(HorizontalAlignment.Right) // 對齊
      .SetWrapText(true);                  // 自動換行
```

其他功能：

- `SetFormula("SUM(A1:A10)")`：設定公式
- `CopyStyleFrom(otherCell)`：複製樣式
- `GetCellValue<T>()`：讀取值

#### 2. 工作簿與工作表管理 (Workbook & Sheet)

方便地管理工作表結構：

```csharp
// 工作表管理
fluent.CloneSheet("Template", "NewReport"); // 複製工作表
fluent.RenameSheet("NewReport", "2024 Report"); // 重新命名
fluent.DeleteSheet("OldData"); // 刪除工作表

// 行列操作
fluent.UseSheet("2024 Report")
      .SetDefaultRowHeight(20) // 預設行高
      .SetRowHeight(0, 30)     // 設定特定行高 (Row 1)
      .SetDefaultColumnWidth(15);
```

#### 3. 圖片操作

```csharp
byte[] imageBytes = File.ReadAllBytes("logo.png");

fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.A, 1)
      .SetPictureOnCell(imageBytes, 200, 100); // 插入圖片並指定寬高
```

#### 4. 高級樣式管理 (Legacy & Dynamic)

除了直接使用 `.Set...` 方法外，也可以使用樣式緩存系統來管理共用樣式，或進行條件格式化。

**註冊共用樣式：**

```csharp
fluent.SetupCellStyle("HeaderStyle", (wb, style) =>
{
    style.SetAlignment(HorizontalAlignment.Center);
    style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
    style.FillPattern = FillPattern.SolidForeground;
});

// 應用樣式
fluent.UseSheet("Sheet1")
      .SetCellPosition(0, 0)
      .SetValue("Title")
      .SetCellStyle("HeaderStyle");
```

**條件格式化 (動態樣式)：**

```csharp
mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithDynamicStyle(item =>
    {
        // 根據資料值返回對應的樣式 Key
        return ((Student)item).Score < 60 ? "FailStyle" : "PassStyle";
    });
```

### 📖 API 概覽

| 用途         | 主要方法                                                                               |
| ------------ | -------------------------------------------------------------------------------------- |
| **Mapping**  | `Map`, `ToColumn`, `WithTitle`, `WithNumberFormat`, `WithBackgroundColor`              |
| **Cell**     | `SetValue`, `SetFormula`, `SetBackgroundColor`, `SetBorder`, `SetFont`, `SetAlignment` |
| **Table**    | `SetTable`, `BuildRows`, `SetAutoFilter`, `FreezeTitleRow`, `AutoSizeColumns`          |
| **Sheet**    | `CloneSheet`, `RenameSheet`, `SetRowHeight`, `SetDefaultColumnWidth`                   |
| **Workbook** | `SaveToPath`, `SaveToStream`, `GetSheetNames`, `DeleteSheet`                           |

---

## English

### 🚀 Features

- ✅ **Fluent API** - Chained method calls for simpler, readable code
- ✅ **Strongly Typed Mapping** - Use `FluentMapping` for type-safe data binding and styling
- ✅ **Direct Styling** - Configure styles directly within Mapping or use FluentCell API
- ✅ **Style Management** - Smart caching to handle duplicate styles and avoid Excel limits
- ✅ **Comprehensive I/O** - Read/Write, Images, Formulas, Merging
- ✅ **Workbook Management** - Clone, Rename, Delete sheets, adjust Row/Column dimensions

### 📦 Installation

```bash
# Via NuGet Package Manager
Install-Package FluentNPOI

# Via .NET CLI
dotnet add package FluentNPOI
```

### 🎯 Quick Start

#### 1. Basic Write

```csharp
using FluentNPOI;
using NPOI.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var fluent = new FluentWorkbook(workbook);

fluent.UseSheet("Sheet1")
      // Position and modify
      .SetCellPosition(0, 1) // Row 1, Col A (0-based)
      .SetValue("Hello World!")
      .SetBackgroundColor(IndexedColors.Yellow) // Styled directly
      .SetFont(isBold: true, fontSize: 14)
      .SetAlignment(HorizontalAlignment.Center);

fluent.SaveToPath("output.xlsx");
```

#### 2. Table Binding with FluentMapping (Recommended)

`FluentMapping` allows you to define both data extraction and visual presentation in one place.

```csharp
var data = new List<Student>
{
    new Student { Name = "Alice", Score = 95.5, Date = DateTime.Now },
    new Student { Name = "Bob", Score = 80.0, Date = DateTime.Now }
};

// Define Mapping & Styles
var mapping = new FluentMapping<Student>();

mapping.Map(x => x.Name)
    .ToColumn(ExcelCol.A)
    .WithTitle("Name")
    .WithAlignment(HorizontalAlignment.Center)
    .WithBackgroundColor(IndexedColors.LightCornflowerBlue);

mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithTitle("Score")
    .WithNumberFormat("0.0") // Set Number Format
    .WithFont(isBold: true);

mapping.Map(x => x.Date)
    .ToColumn(ExcelCol.C)
    .WithTitle("Date")
    .WithNumberFormat("yyyy-mm-dd");

// Write and Enhance
fluent.UseSheet("Report")
      .SetTable(data, mapping)
      .BuildRows()
      .SetAutoFilter() // Add Auto Filter
      .FreezeTitleRow() // Freeze top row
      .AutoSizeColumns(); // Auto-size columns
```

### 📚 Main Features

#### 1. Cell Operations (FluentCell)

FluentCell offers a rich set of chained methods for cell manipulation:

```csharp
fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.C, 5)
      .SetValue(12345.678)
      .SetNumberFormat("#,##0.00")
      .SetBackgroundColor(IndexedColors.Red)
      .SetFont(fontName: "Arial", isBold: true)
      .SetBorder(BorderStyle.Thin)
      .SetAlignment(HorizontalAlignment.Right)
      .SetWrapText(true);
```

Other features:

- `SetFormula("SUM(A1:A10)")`: Set formula
- `CopyStyleFrom(otherCell)`: Copy style
- `GetCellValue<T>()`: Read value

#### 2. Workbook & Sheet Management

Easily manage the structure of your workbook:

```csharp
// Sheet Management
fluent.CloneSheet("Template", "NewReport"); // Clone sheet
fluent.RenameSheet("NewReport", "2024 Report"); // Rename
fluent.DeleteSheet("OldData"); // Delete

// Row & Column Dimensions
fluent.UseSheet("2024 Report")
      .SetDefaultRowHeight(20)
      .SetRowHeight(0, 30) // Set specific row height (Row 1)
      .SetDefaultColumnWidth(15);
```

#### 3. Images

```csharp
byte[] imageBytes = File.ReadAllBytes("logo.png");

fluent.UseSheet("Sheet1")
      .SetCellPosition(ExcelCol.A, 1)
      .SetPictureOnCell(imageBytes, 200, 100); // Insert image with specific size
```

#### 4. Advanced Styling (Legacy & Dynamic)

Besides direct `.Set...` methods, you can use the Style Cache for shared styles or conditional formatting.

**Register Shared Style:**

```csharp
fluent.SetupCellStyle("HeaderStyle", (wb, style) =>
{
    style.SetAlignment(HorizontalAlignment.Center);
    style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
    style.FillPattern = FillPattern.SolidForeground;
});

// Apply Style
fluent.UseSheet("Sheet1")
      .SetCellPosition(0, 0)
      .SetValue("Title")
      .SetCellStyle("HeaderStyle");
```

**Conditional Formatting (Dynamic):**

```csharp
mapping.Map(x => x.Score)
    .ToColumn(ExcelCol.B)
    .WithDynamicStyle(item =>
    {
        // Return style Key based on data
        return ((Student)item).Score < 60 ? "FailStyle" : "PassStyle";
    });
```

### 📖 API Overview

| Area         | Key Methods                                                                            |
| ------------ | -------------------------------------------------------------------------------------- |
| **Mapping**  | `Map`, `ToColumn`, `WithTitle`, `WithNumberFormat`, `WithBackgroundColor`              |
| **Cell**     | `SetValue`, `SetFormula`, `SetBackgroundColor`, `SetBorder`, `SetFont`, `SetAlignment` |
| **Table**    | `SetTable`, `BuildRows`, `SetAutoFilter`, `FreezeTitleRow`, `AutoSizeColumns`          |
| **Sheet**    | `CloneSheet`, `RenameSheet`, `SetRowHeight`, `SetDefaultColumnWidth`                   |
| **Workbook** | `SaveToPath`, `SaveToStream`, `GetSheetNames`, `DeleteSheet`                           |

---

### 🤝 Contribution

Issue and Pull Requests are welcome!

### 📄 License

MIT License - See [LICENSE](LICENSE) file.
