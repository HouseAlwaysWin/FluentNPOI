# FluentNPOI Quick Reference / 快速參考

## 常用操作速查表 / Common Operations Cheat Sheet

### 📝 基本寫入 / Basic Write

```csharp
// 創建工作簿 / Create workbook
var fluent = new FluentWorkbook(new XSSFWorkbook());

// 寫入單個值 / Write single value
fluent.UseSheet("Sheet1")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("Hello");

// 儲存 / Save
fluent.SaveToPath("output.xlsx");
```

### 📖 基本讀取 / Basic Read

```csharp
// 開啟檔案 / Open file
var fluent = new FluentWorkbook(new XSSFWorkbook("file.xlsx"));
var sheet = fluent.UseSheet("Sheet1");

// 讀取值 / Read value
string text = sheet.GetCellValue<string>(ExcelColumns.A, 1);
int number = sheet.GetCellValue<int>(ExcelColumns.B, 1);
DateTime date = sheet.GetCellValue<DateTime>(ExcelColumns.C, 1);
```

### 🎨 樣式設定 / Style Setup

```csharp
// 全局樣式 / Global style
fluent.SetupGlobalCachedCellStyles((wb, style) =>
{
    style.SetAlignment(HorizontalAlignment.Center);
    style.SetBorderAllStyle(BorderStyle.Thin);
});

// 命名樣式 / Named style
fluent.SetupCellStyle("HeaderStyle", (wb, style) =>
{
    style.SetCellFillForegroundColor(IndexedColors.LightBlue);
    style.SetFontInfo(wb, isBold: true);
});
```

### 📊 表格綁定 / Table Binding

```csharp
var data = new List<Person> { /* ... */ };

fluent.UseSheet("People")
    .SetTable(data, ExcelColumns.A, 1)

    .BeginTitleSet("姓名").SetCellStyle("HeaderStyle")
    .BeginBodySet("Name").End()

    .BeginTitleSet("年齡").SetCellStyle("HeaderStyle")
    .BeginBodySet("Age").SetCellType(CellType.Numeric).End()

    .BuildRows();
```

### 🎯 動態樣式 / Dynamic Style

```csharp
.BeginBodySet("Score")
.SetCellStyle(p =>
{
    var score = p.GetRowItem<Student>().Score;
    if (score >= 90)
        return new("HighScore", s => s.SetCellFillForegroundColor("#90EE90"));
    return new("NormalScore", s => s.SetCellFillForegroundColor("#FFFFFF"));
})
.End()
```

### 📋 跨工作表複製樣式 / Copy Style Across Sheets

```csharp
// 從模板工作表複製樣式 / Copy style from template sheet
var templateSheet = fluent.UseSheet("Template");
templateSheet.SetCellPosition(ExcelColumns.A, 1)
    .SetCellStyle("HeaderStyle")
    .SetValue("樣式範本");

// 複製到工作簿級別 / Copy to workbook level
var sheetRef = templateSheet.GetSheet();
fluent.CopyStyleFromSheetCell("copiedStyle", sheetRef, ExcelColumns.A, 1);

// 在其他工作表使用 / Use in other sheets
fluent.UseSheet("Data")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetCellStyle("copiedStyle")
    .SetValue("使用複製的樣式");
```

---

## 常用方法 / Common Methods

### FluentWorkbook

| 方法                                           | 說明           | Example                                                          |
| ---------------------------------------------- | -------------- | ---------------------------------------------------------------- |
| `UseSheet(name)`                               | 使用工作表     | `fluent.UseSheet("Sheet1")`                                      |
| `UseSheet(name, true)`                         | 創建工作表     | `fluent.UseSheet("New", true)`                                   |
| `SetupCellStyle(key, action)`                  | 註冊樣式       | `fluent.SetupCellStyle("MyStyle", ...)`                          |
| `CopyStyleFromSheetCell(key, sheet, col, row)` | 複製單元格樣式 | `fluent.CopyStyleFromSheetCell("key", sheet, ExcelColumns.A, 1)` |
| `SaveToPath(path)`                             | 儲存檔案       | `fluent.SaveToPath("file.xlsx")`                                 |
| `ToStream()`                                   | 輸出串流       | `var stream = fluent.ToStream()`                                 |

### FluentSheet

| 方法                          | 說明     | Example                                    |
| ----------------------------- | -------- | ------------------------------------------ |
| `SetCellPosition(col, row)`   | 設定位置 | `.SetCellPosition(ExcelColumns.A, 1)`      |
| `GetCellValue<T>(col, row)`   | 讀取值   | `.GetCellValue<string>(ExcelColumns.A, 1)` |
| `SetColumnWidth(col, width)`  | 設定欄寬 | `.SetColumnWidth(ExcelColumns.A, 20)`      |
| `SetTable<T>(data, col, row)` | 綁定表格 | `.SetTable(list, ExcelColumns.A, 1)`       |

### FluentCell

| 方法                       | 說明     | Example                      |
| -------------------------- | -------- | ---------------------------- |
| `SetValue(value)`          | 設定值   | `.SetValue("Text")`          |
| `GetValue<T>()`            | 讀取值   | `.GetValue<string>()`        |
| `SetCellStyle(key)`        | 套用樣式 | `.SetCellStyle("MyStyle")`   |
| `SetFormulaValue(formula)` | 設定公式 | `.SetFormulaValue("=A1+B1")` |
| `GetFormula()`             | 讀取公式 | `.GetFormula()`              |

---

## 擴展方法 / Extension Methods

### 樣式相關 / Style Related

```csharp
// 顏色 / Color
style.SetCellFillForegroundColor(255, 0, 0);        // RGB
style.SetCellFillForegroundColor("#FF0000");         // Hex
style.SetCellFillForegroundColor(IndexedColors.Red); // Indexed

// 字型 / Font
style.SetFontInfo(workbook,
    fontFamily: "Arial",
    fontHeight: 12,
    isBold: true,
    color: IndexedColors.Black);

// 邊框 / Border
style.SetBorderAllStyle(BorderStyle.Thin);
style.SetBorderStyle(
    top: BorderStyle.Thick,
    right: BorderStyle.Thin,
    bottom: BorderStyle.Thin,
    left: BorderStyle.Thin);

// 對齊 / Alignment
style.SetAligment(HorizontalAlignment.Center, VerticalAlignment.Center);

// 格式 / Format
style.SetDataFormat(workbook, "yyyy-MM-dd");  // 日期 / Date
style.SetDataFormat(workbook, "#,##0.00");    // 數字 / Number
```

### 工作表相關 / Sheet Related

```csharp
// 欄寬 / Column Width
sheet.SetColumnWidth(ExcelColumns.A, 20);
sheet.SetColumnWidth(ExcelColumns.A, ExcelColumns.E, 15);

// 合併儲存格 / Merge Cells
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.C, 1);        // 橫向 / Horizontal
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.A, 1, 5);    // 縱向 / Vertical
sheet.SetExcelCellMerge(ExcelColumns.A, ExcelColumns.C, 1, 3);    // 區域 / Range

// 取得單元格 / Get Cell
var cell = sheet.GetExcelCell(ExcelColumns.A, 1);
var row = sheet.GetExcelRow(1);
```

---

## 常見模式 / Common Patterns

### 讀取現有檔案並修改 / Read and Modify

```csharp
using var fs = new FileStream("input.xlsx", FileMode.Open);
var fluent = new FluentWorkbook(new XSSFWorkbook(fs));

var sheet = fluent.UseSheet("Sheet1");

// 讀取 / Read
var oldValue = sheet.GetCellValue<string>(ExcelColumns.A, 1);

// 修改 / Modify
sheet.SetCellPosition(ExcelColumns.A, 1)
    .SetValue("New Value");

// 儲存 / Save
fluent.SaveToPath("output.xlsx");
```

### 多工作表操作 / Multi-Sheet Operations

```csharp
var fluent = new FluentWorkbook(new XSSFWorkbook());

// Sheet 1
fluent.UseSheet("Summary")
    .SetCellPosition(ExcelColumns.A, 1)
    .SetValue("總計");

// Sheet 2
fluent.UseSheet("Details", true)
    .SetTable(data, ExcelColumns.A, 1)
    .BuildRows();

// Sheet 3
fluent.UseSheet(0)
    .SetCellPosition(ExcelColumns.B, 1)
    .SetValue("Updated");

fluent.SaveToPath("multi-sheet.xlsx");
```

### 條件格式 / Conditional Formatting

```csharp
.SetTable(salesData, ExcelColumns.A, 1)

.BeginTitleSet("銷售額")
.BeginBodySet("Amount")
.SetCellStyle(p =>
{
    var amount = p.GetRowItem<Sale>().Amount;

    if (amount > 10000)
        return new("High", s => s.SetCellFillForegroundColor("#90EE90"));
    else if (amount > 5000)
        return new("Medium", s => s.SetCellFillForegroundColor("#FFFFE0"));
    else
        return new("Low", s => s.SetCellFillForegroundColor("#FFB6C1"));
})
.End()

.BuildRows();
```

### DataTable 綁定 / DataTable Binding

```csharp
DataTable dt = GetDataTable();

fluent.UseSheet("DataSheet")
    .SetTable<DataRow>(dt.Rows.Cast<DataRow>(), ExcelColumns.A, 1)

    .BeginTitleSet("欄位1")
    .BeginBodySet("Column1").End()

    .BeginTitleSet("欄位2")
    .BeginBodySet("Column2")
    .SetCellStyle(p =>
    {
        var row = p.RowItem as DataRow;
        var value = row["Column2"].ToString();

        if (value == "特殊")
            return new("Special", s => s.SetCellFillForegroundColor("#FFFF00"));
        return new("Normal", s => { });
    })
    .End()

    .BuildRows();
```

---

## 資料類型對應 / Data Type Mapping

| C# Type                      | Excel Type     | 注意事項 / Notes                |
| ---------------------------- | -------------- | ------------------------------- |
| `string`                     | Text           | 自動處理 / Auto                 |
| `int`, `long`                | Numeric        | 自動轉換 / Auto convert         |
| `double`, `decimal`, `float` | Numeric        | 自動轉換 / Auto convert         |
| `bool`                       | Boolean        | 自動處理 / Auto                 |
| `DateTime`                   | Numeric (Date) | 需要日期格式 / Need date format |
| `DBNull`, `null`             | Blank          | 空白單元格 / Empty cell         |

---

## 效能提示 / Performance Tips

### ✅ 好的做法 / Good Practices

```csharp
// 1. 使用樣式緩存 / Use style caching
fluent.SetupCellStyle("MyStyle", (wb, s) => { /* ... */ });

// 2. 批次操作 / Batch operations
fluent.UseSheet("Data")
    .SetTable(largeList, ExcelColumns.A, 1)
    .BuildRows();

// 3. 重用 Key / Reuse keys
return new CellStyleConfig("consistent-key", style => { /* ... */ });
```

### ❌ 避免的做法 / Bad Practices

```csharp
// 1. 每次創建新樣式 / Creating new style every time
return new CellStyleConfig("", style => { /* ... */ }); // Empty key!

// 2. 逐個單元格操作 / Cell by cell operations
for (int i = 0; i < 10000; i++)
{
    sheet.SetCellPosition(ExcelColumns.A, i).SetValue(data[i]);
}

// 3. 動態生成唯一 Key / Dynamic unique keys
return new CellStyleConfig($"style-{Guid.NewGuid()}", style => { /* ... */ });
```

---

## 疑難排解 / Troubleshooting

### 問題：樣式超過 64000 限制

**解決方案**：使用一致的 Key

```csharp
// ❌ 錯誤 / Wrong
.SetCellStyle(p => new("", s => { })); // 每次創建新樣式 / Creates new style

// ✅ 正確 / Correct
.SetCellStyle(p => new("my-key", s => { })); // 重用樣式 / Reuses style
```

### 問題：日期顯示為數字

**解決方案**：設定日期格式

```csharp
fluent.SetupCellStyle("DateFormat", (wb, style) =>
{
    style.SetDataFormat(wb, "yyyy-MM-dd");
});

sheet.SetCellPosition(ExcelColumns.A, 1)
    .SetValue(DateTime.Now)
    .SetCellStyle("DateFormat");
```

### 問題：讀取值類型不正確

**解決方案**：使用泛型指定類型

```csharp
// 自動判斷 / Auto detect
var value = sheet.GetCellValue(ExcelColumns.A, 1);

// 指定類型 / Specify type
var text = sheet.GetCellValue<string>(ExcelColumns.A, 1);
var number = sheet.GetCellValue<double>(ExcelColumns.A, 1);
```

---

## 更多資源 / More Resources

- 📖 [完整文檔 / Full Documentation](README.md)
- 💻 [範例程式 / Examples](FluentNPOIConsoleExample/Program.cs)
- 🧪 [單元測試 / Unit Tests](FluentNPOIUnitTest/UnitTest1.cs)
- 🤝 [貢獻指南 / Contributing](CONTRIBUTING.md)
- 📝 [變更記錄 / Changelog](CHANGELOG.md)
