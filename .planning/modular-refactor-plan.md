# FluentNPOI 模組化拆分計畫

## 背景
FluentNPOI 目前依賴太多套件（NPOI、QuestPDF、ExcelDataReader、System.Drawing.Common），計畫加入 ScottPlot 圖表功能。需要拆分成多個獨立套件，讓使用者按需安裝。

## 目標架構

```
FluentNPOI/
├── FluentNPOI/                  ← 核心套件（僅 NPOI + Microsoft.CSharp）
├── FluentNPOI.Pdf/              ← QuestPDF 整合
├── FluentNPOI.Streaming/        ← ExcelDataReader 整合  
├── FluentNPOI.Charts/           ← ScottPlot 整合（新）
└── FluentNPOI.All/              ← 元套件，依賴全部子套件
```

## 套件依賴關係

| 套件 | 依賴 | 功能 |
|------|------|------|
| **FluentNPOI** | NPOI | 基本 Excel 讀寫、樣式、圖片 |
| **FluentNPOI.Pdf** | FluentNPOI + QuestPDF | PDF 匯出 |
| **FluentNPOI.Streaming** | FluentNPOI + ExcelDataReader | 大檔案串流讀取 |
| **FluentNPOI.Charts** | FluentNPOI + ScottPlot | 圖表生成並嵌入 Excel |
| **FluentNPOI.All** | 上述全部 | 完整功能元套件 |

## 實作步驟

### Phase 1: 準備工作 ✅
- [x] 建立 solution 結構，新增子專案資料夾
- [x] 定義共用的 NuGet metadata（作者、License、Repository URL）

### Phase 2: 核心套件重構 (進行中)
- [x] 複製 Pdf/ 程式碼到 FluentNPOI.Pdf
- [x] 複製 Streaming/ 程式碼到 FluentNPOI.Streaming
- [ ] 從 FluentNPOI 移除 QuestPDF、ExcelDataReader 依賴
- [ ] 刪除核心套件的 Pdf/、Streaming/ 資料夾
- [ ] 定義擴展點介面（如 IExcelExporter）

### Phase 3: 建立擴展套件
- [ ] FluentNPOI.Pdf - 移動 Pdf/ 資料夾內容
- [ ] FluentNPOI.Streaming - 移動 Streaming/ 資料夾內容
- [ ] FluentNPOI.Charts - 新建 ScottPlot 整合

### Phase 4: 元套件與測試
- [ ] 建立 FluentNPOI.All 元套件
- [ ] 更新單元測試專案
- [ ] 更新 CI/CD workflow 支援多套件發布

### Phase 5: 文件與發布
- [ ] 更新 README 說明模組化安裝方式
- [ ] 發布所有套件到 NuGet

## 使用範例

```csharp
// 只需基本功能
dotnet add package FluentNPOI

// 需要 PDF 匯出
dotnet add package FluentNPOI.Pdf

// 需要圖表
dotnet add package FluentNPOI.Charts

// 全部功能
dotnet add package FluentNPOI.All
```

## 注意事項
- 版本號需保持同步
- 核心套件 API 變更會影響所有擴展套件
- 考慮使用 Directory.Build.props 統一管理版本
