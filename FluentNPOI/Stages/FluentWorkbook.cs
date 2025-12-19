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
    /// Workbook operation class
    /// </summary>
    public class FluentWorkbook : FluentWorkbookBase
    {
        private ISheet _currentSheet;

        /// <summary>
        /// Initialize FluentWorkbook instance
        /// </summary>
        /// <param name="workbook">NPOI Workbook object</param>
        public FluentWorkbook(IWorkbook workbook)
            : base(workbook, new Dictionary<string, ICellStyle>())
        {
        }

        /// <summary>
        /// Read Excel file
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook ReadExcelFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel file not found.", filePath);

            _currentSheet = null;

            string ext = Path.GetExtension(filePath)?.ToLowerInvariant();

            // Open with Read mode and read into memory immediately, release file lock after reading
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (ext == ".xls")
                {
                    _workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    // Default use XSSF, supports .xlsx/.xlsm
                    _workbook = new XSSFWorkbook(fs);
                }
            }

            // Select first sheet by default
            if (_workbook.NumberOfSheets > 0)
            {
                _currentSheet = _workbook.GetSheetAt(0);
            }

            return this;
        }

        /// <summary>
        /// Copy style from cell in specified sheet
        /// </summary>
        /// <param name="cellStyleKey">Style cache key</param>
        /// <param name="sheet">Source sheet</param>
        /// <param name="col">Column position</param>
        /// <param name="rowIndex">Row index (1-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
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
        /// Set global cached cell style (clears all existing styles)
        /// </summary>
        /// <param name="styles">Style configuration action</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetupGlobalCachedCellStyles(Action<IWorkbook, ICellStyle> styles)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            styles(_workbook, newCellStyle);
            _cellStylesCached.Clear();
            _cellStylesCached.Add("global", newCellStyle);
            return this;
        }

        /// <summary>
        /// Set and cache cell style
        /// </summary>
        /// <param name="cellStyleKey">Style cache key</param>
        /// <param name="styles">Style configuration action</param>
        /// <param name="inheritFrom">Optional, inherited parent style key. If specified, copies all properties from parent first, then applies custom changes</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetupCellStyle(string cellStyleKey, Action<IWorkbook, ICellStyle> styles, string inheritFrom = null)
        {
            ICellStyle newCellStyle = _workbook.CreateCellStyle();

            // If inheritance is specified, copy all properties from parent style first
            if (!string.IsNullOrEmpty(inheritFrom) && _cellStylesCached.TryGetValue(inheritFrom, out var parentStyle))
            {
                newCellStyle.CloneStyleFrom(parentStyle);
            }

            // Apply custom changes (will override properties from parent style)
            styles(_workbook, newCellStyle);
            _cellStylesCached.Add(cellStyleKey, newCellStyle);
            return this;
        }


        /// <summary>
        /// Use specified sheet
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <param name="createIfMissing">Create if missing</param>
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
        /// Use specified sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public FluentSheet UseSheet(ISheet sheet)
        {
            _currentSheet = sheet;
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// Use sheet at specified index
        /// </summary>
        /// <param name="index">Index</param>
        /// <param name="createIfMissing">Create if missing</param>
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

        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SaveToFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            // Ensure directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(fs);
            }
            return this;
        }

        /// <summary>
        /// Save to stream
        /// </summary>
        /// <param name="stream">Target stream</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SaveToStream(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            _workbook.Write(stream);
            return this;
        }

        /// <summary>
        /// Get all sheet names
        /// </summary>
        /// <returns>List of sheet names</returns>
        public List<string> GetSheetNames()
        {
            var names = new List<string>();
            for (int i = 0; i < _workbook.NumberOfSheets; i++)
            {
                names.Add(_workbook.GetSheetName(i));
            }
            return names;
        }

        /// <summary>
        /// Get sheet count
        /// </summary>
        public int SheetCount => _workbook.NumberOfSheets;

        /// <summary>
        /// Delete sheet by name
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook DeleteSheet(string sheetName)
        {
            int index = _workbook.GetSheetIndex(sheetName);
            if (index >= 0)
            {
                _workbook.RemoveSheetAt(index);
            }
            return this;
        }

        /// <summary>
        /// Delete sheet by index
        /// </summary>
        /// <param name="index">Sheet index (0-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook DeleteSheetAt(int index)
        {
            if (index >= 0 && index < _workbook.NumberOfSheets)
            {
                _workbook.RemoveSheetAt(index);
            }
            return this;
        }

        /// <summary>
        /// Clone sheet
        /// </summary>
        /// <param name="sourceSheetName">Source sheet name</param>
        /// <param name="newSheetName">New sheet name</param>
        /// <returns>New FluentSheet instance</returns>
        public FluentSheet CloneSheet(string sourceSheetName, string newSheetName)
        {
            int sourceIndex = _workbook.GetSheetIndex(sourceSheetName);
            if (sourceIndex < 0)
                throw new ArgumentException($"Sheet '{sourceSheetName}' not found.", nameof(sourceSheetName));

            ISheet clonedSheet = _workbook.CloneSheet(sourceIndex);
            int clonedIndex = _workbook.GetSheetIndex(clonedSheet);
            _workbook.SetSheetName(clonedIndex, newSheetName);
            _currentSheet = clonedSheet;
            return new FluentSheet(_workbook, _currentSheet, _cellStylesCached);
        }

        /// <summary>
        /// Rename sheet
        /// </summary>
        /// <param name="oldName">Old name</param>
        /// <param name="newName">New name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook RenameSheet(string oldName, string newName)
        {
            int index = _workbook.GetSheetIndex(oldName);
            if (index >= 0)
            {
                _workbook.SetSheetName(index, newName);
            }
            return this;
        }

        /// <summary>
        /// Set active sheet (sheet shown when opening Excel)
        /// </summary>
        /// <param name="index">Sheet index (0-based)</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetActiveSheet(int index)
        {
            if (index >= 0 && index < _workbook.NumberOfSheets)
            {
                _workbook.SetActiveSheet(index);
            }
            return this;
        }

        /// <summary>
        /// Set active sheet (sheet shown when opening Excel)
        /// </summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>FluentWorkbook instance, supports method chaining</returns>
        public FluentWorkbook SetActiveSheet(string sheetName)
        {
            int index = _workbook.GetSheetIndex(sheetName);
            if (index >= 0)
            {
                _workbook.SetActiveSheet(index);
            }
            return this;
        }
    }
}

