using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using FluentNPOI.Base;
using System;
using System.Collections.Generic;
using System.IO;
using FluentNPOI.Models;

namespace FluentNPOI.Stages
{
    /// <summary>
    /// 工作簿操作類
    /// </summary>
    public class FluentWorkbook : FluentWorkbookBase
    {
        private ISheet _currentSheet;

        /// <summary>
        /// 初始化 FluentWorkbook 实例
        /// </summary>
        /// <param name="workbook">NPOI 工作簿对象</param>
        public FluentWorkbook(IWorkbook workbook)
            : base(workbook, new Dictionary<string, ICellStyle>())
        {
        }

        /// <summary>
        /// 读取 Excel 文件
        /// </summary>
        /// <param name="filePath">Excel 文件路径</param>
        /// <returns>FluentWorkbook 实例，支持链式调用</returns>
        public FluentWorkbook ReadExcelFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel file not found.", filePath);

            _currentSheet = null;

            string ext = Path.GetExtension(filePath)?.ToLowerInvariant();

            // 以讀取模式開啟並立即讀入記憶體，讀完即釋放檔案鎖
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (ext == ".xls")
                {
                    _workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    // 預設使用 XSSF，支援 .xlsx/.xlsm
                    _workbook = new XSSFWorkbook(fs);
                }
            }

            // 預設選取第一個工作表
            if (_workbook.NumberOfSheets > 0)
            {
                _currentSheet = _workbook.GetSheetAt(0);
            }

            return this;
        }

        /// <summary>
        /// 从指定工作表的单元格复制样式
        /// </summary>
        /// <param name="cellStyleKey">样式缓存键名</param>
        /// <param name="sheet">源工作表</param>
        /// <param name="col">列位置</param>
        /// <param name="rowIndex">行位置（1-based）</param>
        /// <returns>FluentWorkbook 实例，支持链式调用</returns>
        public FluentWorkbook CopyStyleFromSheetCell(string cellStyleKey, ISheet sheet, ExcelCol col, int rowIndex)
        {
            ICell cell = sheet.GetExcelCell(col, rowIndex);
            if (cell != null && cell.CellStyle != null && !_cellStylesCached.ContainsKey(cellStyleKey))
            {
                ICellStyle newCellStyle = _workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(cell.CellStyle);
                _cellStylesCached.Add(cellStyleKey, newCellStyle);
            }
            return this;
        }

        /// <summary>
        /// 设置全局缓存的单元格样式（会清除所有现有样式）
        /// </summary>
        /// <param name="styles">样式设置函数</param>
        /// <returns>FluentWorkbook 实例，支持链式调用</returns>
        public FluentWorkbook SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            styles(_workbook, newCellStyle);
            _cellStylesCached.Clear();
            _cellStylesCached.Add("global", newCellStyle);
            return this;
        }

        /// <summary>
        /// 设置并缓存单元格样式
        /// </summary>
        /// <param name="cellStyleKey">样式缓存键名</param>
        /// <param name="styles">样式设置函数</param>
        /// <param name="inheritFrom">可選，繼承的父樣式鍵名。若指定，會先複製父樣式的所有屬性，再套用自訂修改</param>
        /// <returns>FluentWorkbook 实例，支持链式调用</returns>
        public FluentWorkbook SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles, string inheritFrom = null)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();

            // 如果指定繼承，先從父樣式複製所有屬性
            if (!string.IsNullOrEmpty(inheritFrom) && _cellStylesCached.TryGetValue(inheritFrom, out var parentStyle))
            {
                newCellStyle.CloneStyleFrom(parentStyle);
            }

            // 套用自訂修改（會覆寫父樣式的對應屬性）
            styles(_workbook, newCellStyle);
            _cellStylesCached.Add(cellStyleKey, newCellStyle);
            return this;
        }


        /// <summary>
        /// 使用指定工作表
        /// </summary>
        /// <param name="sheetName">工作表名稱</param>
        /// <param name="createIfMissing">如果不存在是否建立</param>
        /// <returns></returns>
        public FluentSheet UseSheet(string sheetName, bool createIfMissing = true)
        {
            _currentSheet = _workbook.GetSheet(sheetName);
            if (_currentSheet == null && createIfMissing)
            {
                _currentSheet = _workbook.CreateSheet(sheetName);
            }
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// 使用指定工作表
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public FluentSheet UseSheet(ISheet sheet)
        {
            _currentSheet = sheet;
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// 使用指定索引的工作表
        /// </summary>
        /// <param name="index">索引</param>
        /// <param name="createIfMissing">如果不存在是否建立</param>
        /// <returns></returns>
        public FluentSheet UseSheetAt(int index, bool createIfMissing = false)
        {
            _currentSheet = _workbook.GetSheetAt(index);
            if (_currentSheet == null && createIfMissing)
            {
                _currentSheet = _workbook.CreateSheet();
            }
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }
    }
}

