using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOIPlus.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace NPOIPlus
{

	public class NPOIWorkbook
	{
		/// <summary>
		/// 儲存格類型列舉
		/// </summary>
		private enum CellTypeEnum
		{
			Int,
			Double,
			DateTime,
			String
		}

		public IWorkbook Workbook { get; set; }

		public DefaultType<int> SetDefaultIntCellValue = (value) => value;
		public DefaultType<double> SetDefaultDoubleCellValue = (value) => value;
		public DefaultType<bool> SetDefaultBoolCellValue = (value) => value;
		public DefaultType<string> SetDefaultStringCellValue = (value) => value;
		public DefaultType<DateTime> SetDefaultDateTimeCellValue = (value) => value;


		public Action<ICellStyle> SetGlobalCellStyle = (style) => { };
		public Action<ICellStyle> SetDefaultIntCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultDoubleCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultNumberCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultBoolCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultStringCellStyle = (value) => { };

		public Action<ICellStyle> SetDefaultDateTimeCellStyle = (value) => { };

		/// <summary>
		/// 廠區Excel通用格式
		/// </summary>
		public void SetDlpDefaultExcelStyle()
		{
			IWorkbook workbook = this.Workbook;

			SetGlobalCellStyle = (style) =>
			{
				style.SetBorderAllStyle(BorderStyle.Thin);
			};

			SetDefaultNumberCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "#,##0");
				style.SetAligment(HorizontalAlignment.Right);
			};
			SetDefaultStringCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "#,##0");
				style.SetAligment(HorizontalAlignment.Left);

			};
			SetDefaultDateTimeCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "yyyy/MM/DD");
				style.SetAligment(HorizontalAlignment.Center);
			};
		}

		/// <summary>
		/// 廠區Excel通用格式
		/// </summary>
		public void SetDlpDefaultExcelStyleNoBorder()
		{
			IWorkbook workbook = this.Workbook;

			SetDefaultNumberCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "#,##0");
				style.SetAligment(HorizontalAlignment.Right);
			};
			SetDefaultStringCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "#,##0");
				style.SetAligment(HorizontalAlignment.Left);

			};
			SetDefaultDateTimeCellStyle = (style) =>
			{
				style.SetDataFormat(workbook, "yyyy/MM/DD");
				style.SetAligment(HorizontalAlignment.Center);
			};
		}

		// 樣式快取
		private Dictionary<string, ICellStyle> _cellStylesCached = new Dictionary<string, ICellStyle>();

		public void SetupGlobalCachedCellStyles(ISheet sheet, List<StyleCachedParam> styles)
		{
			foreach (var item in styles)
			{
				AddCellStyles(sheet, item.Key, item.StyleAction);
			}
		}

		public string AddCellStyles(ISheet sheet, string key, Action<ICellStyle> style)
		{
			// string cachedKey = SetGlobalStyleKeyBasedOnType("", $"{sheet.SheetName}_{key}", cellType);
			string cachedKey = $"{sheet.SheetName}_{key}";
			if (!_cellStylesCached.ContainsKey(cachedKey))
			{
				ICellStyle newCellStyle = Workbook.CreateCellStyle();
				style(newCellStyle);
				_cellStylesCached.Add(cachedKey, newCellStyle);
			}
			return key;
		}

		/// <summary>
		/// 複製特定欄位的Style
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="col"></param>
		/// <param name="rowIndex"></param>
		/// <param name="key"></param>
		/// <param name="cellType"></param>
		public string CopyStyleFromCell(ISheet sheet, ExcelColumns col, int rowIndex)
		{
			string key = $"{col}{rowIndex}";
			ICell cell = sheet.GetExcelCell(col, rowIndex);
			AddCellStyles(sheet, key, (style) =>
			{
				style.CloneStyleFrom(cell.CellStyle);
			});
			return key;
		}

		public List<string> GetCurrentCellStylesCached()
		{
			return _cellStylesCached.Keys.ToList();
		}
		public NPOIWorkbook(IWorkbook workbook)
		{
			Workbook = workbook;
		}

		/// <summary>
		/// 判斷儲存格值的實際類型
		/// </summary>
		/// <param name="cellValue">儲存格值</param>
		/// <param name="cellType">明確指定的類型</param>
		/// <returns>CellTypeEnum 列舉值</returns>
		private CellTypeEnum DetermineCellType(object cellValue, Type cellType = null)
		{
			// 明確的類型參數優先級最高
			if (cellType != null)
			{
				if (cellType == typeof(int)) return CellTypeEnum.Int;
				if (cellType == typeof(double) || cellType == typeof(float)) return CellTypeEnum.Double;
				if (cellType == typeof(DateTime)) return CellTypeEnum.DateTime;
				if (cellType == typeof(string)) return CellTypeEnum.String;
			}

			// 如果無值，回傳字串型別
			if (cellValue == null || cellValue == DBNull.Value) return CellTypeEnum.String;

			var stringValue = cellValue.ToString();

			// 嘗試按優先級判斷型別
			if (int.TryParse(stringValue, out _)) return CellTypeEnum.Int;
			if (double.TryParse(stringValue, out _)) return CellTypeEnum.Double;
			if (DateTime.TryParse(stringValue, out _)) return CellTypeEnum.DateTime;

			return CellTypeEnum.String;
		}

		/// <summary>
		/// 安全地轉換值為字串
		/// </summary>
		private string SafeToString(object obj)
		{
			return obj?.ToString() ?? string.Empty;
		}

		/// <summary>
		/// 驗證與標準化列號
		/// </summary>
		private int NormalizeRow(int row)
		{
			return row < 1 ? 1 : row;
		}

		private void SetCellStyleBasedOnType(object cellValue, ICellStyle style, Type cellType = null)
		{
			if (cellValue == null || cellValue == DBNull.Value) return;

			var typeEnum = DetermineCellType(cellValue, cellType);
			
			switch (typeEnum)
			{
				case CellTypeEnum.Int:
					SetDefaultNumberCellStyle?.Invoke(style);
					SetDefaultIntCellStyle?.Invoke(style);
					break;
				case CellTypeEnum.Double:
					SetDefaultNumberCellStyle?.Invoke(style);
					SetDefaultDoubleCellStyle?.Invoke(style);
					break;
				case CellTypeEnum.DateTime:
					SetDefaultDateTimeCellStyle?.Invoke(style);
					break;
				case CellTypeEnum.String:
				default:
					SetDefaultStringCellStyle?.Invoke(style);
					break;
			}
		}

		/// <summary>
		/// 設定欄位A~N自動寬度
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		public void SetRangeAutoSizeColumns(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol)
		{
			int startColIndex = (int)startCol;
			int endColIndex = (int)endCol;
			for (int j = startColIndex; j <= endColIndex; j++)
			{
				sheet.AutoSizeColumn(j);
			}
		}


		/// <summary>
		///  設定範圍樣式(後蓋前)
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <param name="style"></param>
		/// <param name="overrideStyle">是否複寫原本的樣式</param>
		public void SetRangeCellStyle(ISheet sheet, ExcelColumns col, int row, Action<ICellStyle> style, bool overrideStyle = false, object defaultValue = null)
		{
			SetRangeCellStyle(sheet, col, col, row, row, style, overrideStyle, defaultValue: defaultValue);
		}

		/// <summary>
		///  設定範圍樣式(後蓋前)
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="col"></param>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="style"></param>
		/// <param name="overrideStyle">是否複寫原本的樣式</param>
		public void SetRangeCellStyle(ISheet sheet, ExcelColumns col, int startRow, int endRow, Action<ICellStyle> style, bool overrideStyle = false, object defaultValue = null)
		{
			SetRangeCellStyle(sheet, col, col, startRow, endRow, style, overrideStyle, defaultValue: defaultValue);
		}

		/// <summary>
		///  設定範圍樣式(後蓋前)
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		/// <param name="row"></param>
		/// <param name="style"></param>
		/// <param name="overrideStyle">是否複寫原本的樣式</param>
		public void SetRangeCellStyle(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol, int row, Action<ICellStyle> style, bool overrideStyle = false, object defaultValue = null)
		{
			SetRangeCellStyle(sheet, startCol, endCol, row, row, style, overrideStyle, defaultValue: defaultValue);
		}


		/// <summary>
		/// 
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="cellType"></param>
		/// <param name="cellStyleKey"></param>
		public void SetRangeCellStyle(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol, int startRow, int endRow, string cellStyleKey, bool mergeCell = false, bool overrideStyle = false, object defaultValue = null)
		{
			if (mergeCell)
			{
				sheet.SetExcelCellMerge(startCol, endCol, startRow, endRow);
			}
			// string cachedKey = SetGlobalStyleKeyBasedOnType("", $"{sheet.SheetName}_{cellStyleKey}", cellType);
			string cachedKey = $"{sheet.SheetName}_{cellStyleKey}";
			if (!_cellStylesCached.ContainsKey(cachedKey)) throw new Exception($"{cachedKey}樣式不存在");
			SetRangeCellStyle(sheet, startCol, endCol, startRow, endRow, null, overrideStyle, cachedKey, defaultValue: defaultValue);
		}

		public void SetRangeCellStyle(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol, int row, string cellStyleKey, bool mergeCell = false, bool overrideStyle = false, object defaultValue = null)
		{
			// SetRangeCellStyle(sheet, startCol, endCol, row, row, cellType, cellStyleKey, mergeCell);
			SetRangeCellStyle(sheet, startCol, endCol, row, row, cellStyleKey, mergeCell, overrideStyle, defaultValue: defaultValue);
		}

		/// <summary>
		///  設定範圍樣式(後蓋前)
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="style"></param>
		/// <param name="overrideStyle">是否複寫原本的樣式</param>
		public void SetRangeCellStyle(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol, int startRow, int endRow, Action<ICellStyle> style, bool overrideStyle = false, string cellStyleKey = "", object defaultValue = null)
		{
			int startColIndex = (int)startCol;
			int endColIndex = (int)endCol;
			startRow = NormalizeRow(startRow);
			endRow = NormalizeRow(endRow);
			string styleCachedKey = $"SetRangeCellStyle_{startCol}_{endCol}_{startRow}_{endRow}";
			if (!string.IsNullOrWhiteSpace(cellStyleKey))
			{
				styleCachedKey = cellStyleKey;
			}

			for (int i = startRow; i <= endRow; i++)
			{
				int rowIndex = i - 1;
				IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
				if (row != null)
				{
					for (int j = startColIndex; j <= endColIndex; j++)
					{
						ICell cell = row.GetCell(j);

						if (cell == null || overrideStyle || cell?.CellType == CellType.Blank)
						{

							cell = cell ?? row.CreateCell(j);
							if (defaultValue != null)
							{
								SetCellValueBasedOnType(cell, defaultValue);
							}
							var styleCachedDict = _cellStylesCached;
							if (styleCachedDict.ContainsKey(styleCachedKey))
							{
								cell.CellStyle = styleCachedDict[styleCachedKey]; // 直接使用已存在的樣式
							}
							else
							{
								ICellStyle newCellStyle = Workbook.CreateCellStyle();
								style(newCellStyle);
								cell.CellStyle = newCellStyle;
								styleCachedDict.Add(styleCachedKey, newCellStyle);
							}
						}

					}
				}
			}
		}

		private void SetCellValueBasedOnType(ICell cell, object cellValue, CellValueActionType valueAction = null,
			ExcelColumns colnum = 0, int rownum = 1, Type cellType = null)
		{
			if (cell == null || cellValue == null || cellValue == DBNull.Value)
				return;

			var newCellValue = valueAction?.Invoke(cell, cellValue, colnum, rownum) ?? cellValue;
			if (newCellValue == null || newCellValue == DBNull.Value)
				return;

			var typeEnum = DetermineCellType(cellValue, cellType);
			var stringValue = SafeToString(newCellValue);

			switch (typeEnum)
			{
				case CellTypeEnum.Int:
					if (int.TryParse(stringValue, out var intVal))
					{
						var processedValue = SetDefaultIntCellValue(intVal);
						cell.SetCellValue(processedValue);
					}
					break;
				case CellTypeEnum.Double:
					if (double.TryParse(stringValue, out var doubleVal))
					{
						var processedValue = SetDefaultDoubleCellValue(doubleVal);
						cell.SetCellValue(processedValue);
					}
					break;
				case CellTypeEnum.DateTime:
					if (DateTime.TryParse(stringValue, out var dateVal))
					{
						var processedValue = SetDefaultDateTimeCellValue(dateVal);
						cell.SetCellValue(processedValue);
					}
					break;
				case CellTypeEnum.String:
				default:
					var processedString = SetDefaultStringCellValue(stringValue);
					cell.SetCellValue(processedString);
					break;
			}
		}

		private string SetGlobalStyleKeyBasedOnType(object cellValue, string key, Type cellType)
		{
			var typeEnum = DetermineCellType(cellValue, cellType);
			
			switch (typeEnum)
			{
				case CellTypeEnum.Int:
					return $"Int_{key}";
				case CellTypeEnum.Double:
                    return $"Double_{key}";
				case CellTypeEnum.DateTime:
                    return $"DateTime_{key}";
				case CellTypeEnum.String:
                    return $"String_{key}";
				default:
                    return $"String_{key}";
			}
		}

		private void SetCellStyle(string cachedKey, ICell cell, object cellValue, Action<ICellStyle> colStyle = null,
			Action<ICellStyle> rowStyle = null, ExcelColumns colnum = 0, int rownum = 1, Type cellType = null, string cellStyleKey = null)
		{
			string key = SetGlobalStyleKeyBasedOnType(cellValue, "GlobalStyle", cellType);

			// 自定義的key值
			if (!string.IsNullOrWhiteSpace(cellStyleKey))
			{
				// key = SetGlobalStyleKeyBasedOnType(cellValue, $"{cachedKey}_{cellStyleKey}", cellType);
				key = $"{cachedKey}_{cellStyleKey}";
			}
			// 設定整行的欄位 例如 A行
			else if (colStyle != null)
			{
				key = SetGlobalStyleKeyBasedOnType(cellValue, $"{cachedKey}_ColStyle_{colnum}", cellType);
			}
			// 設定整排的Style 例如 A到L欄位
			else if (colStyle == null && rowStyle != null)
			{
				key = SetGlobalStyleKeyBasedOnType(cellValue, $"{cachedKey}_RowStyle_{rownum}", cellType);
			}

			// 檢查是否已有樣式
			if (_cellStylesCached.ContainsKey(key))
			{
				cell.CellStyle = _cellStylesCached[key]; // 直接使用已存在的樣式
			}
			else
			{
				ICellStyle newCellStyle = Workbook.CreateCellStyle();
				SetGlobalCellStyle(newCellStyle);
				SetCellStyleBasedOnType(cellValue, newCellStyle, cellType);
				rowStyle?.Invoke(newCellStyle);
				colStyle?.Invoke(newCellStyle);

				cell.CellStyle = newCellStyle;

				_cellStylesCached.Add(key, newCellStyle);
			}
		}


		/// <summary>
		/// 設定單一個欄位
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellValue"></param>
		/// <param name="colnum"></param>
		/// <param name="rownum"></param>
		/// <param name="param"></param>
		/// <exception cref="Exception"></exception>
		public void SetExcelCell<T>(ISheet sheet, T cellValue, ExcelColumns colnum, int rownum,
			Action<ICellStyle> style = null, bool? isFormula = null, Type cellType = null, string cellStyleKey = null)
		{
			rownum = NormalizeRow(rownum);
			var key = sheet.SheetName;

			SetExcelCell(sheet, key, new DataTable(), 0, "", colnum, rownum, cellValue, style, null, null, isFormula, cellType, cellStyleKey);
		}

		/// <summary>
		/// 設定單一個欄位
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="tableIndex"></param>
		/// <param name="tableColName"></param>
		/// <param name="column"></param>
		/// <param name="rownum"></param>
		/// <param name="colStyle"></param>
		/// <param name="isFormula"></param>
		/// <param name="cellType"></param>
		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName,
			ExcelColumns column, int rownum = 1, Action<ICellStyle> colStyle = null, bool? isFormula = null, Type cellType = null)
		{
			SetExcelCell(sheet, dataTable, tableIndex, tableColName, column, rownum, null, colStyle, null, null,
				isFormula, cellType);
		}

		private void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName,
			ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> colStyle = null,
			Action<ICellStyle> rowStyle = null, CellValueActionType cellValueAction = null, bool? isFormula = false, Type cellType = null)
		{
			var key = sheet.SheetName;

			SetExcelCell(sheet, key, dataTable, tableIndex, tableColName, colnum, rownum, cellValue, colStyle, rowStyle,
				cellValueAction, isFormula, cellType);
		}

		private void SetExcelCell(ISheet sheet, string groupKey, DataTable dataTable, int tableIndex,
			string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null,
			Action<ICellStyle> colStyle = null, Action<ICellStyle> rowStyle = null,
			CellValueActionType cellValueAction = null, bool? isFormula = false, Type cellType = null, string cellStyleKey = null)
		{
			rownum = NormalizeRow(rownum);
			int zeroBaseIndex = rownum - 1;
			IRow row = sheet.GetRow(zeroBaseIndex) ?? sheet.CreateRow(zeroBaseIndex);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = dataTable.Columns.Contains(tableColName)
				? dataTable.Rows?[tableIndex]?[tableColName]
				: cellValue;

			SetCellStyle(groupKey, cell, newValue, colStyle, rowStyle, colnum, rownum, cellType, cellStyleKey);

			// 設定CellValue
			if (isFormula.HasValue)
			{
				if (isFormula.Value)
				{
					object newCellValue = cellValueAction?.Invoke(cell, cellValue, colnum, rownum) ??
										  cellValue?.ToString();
					cell.SetCellFormula(SafeToString(newCellValue));
					return;
				}
			}
			SetCellValueBasedOnType(cell, newValue, cellValueAction, colnum, rownum, cellType);
		}

		/// <summary>
		/// 設定一行Row
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="tableIndex"></param>
		/// <param name="param"></param>
		/// <param name="startColnum"></param>
		/// <param name="rownum"></param>
		/// <param name="rowStyle"></param>
		/// <param name="isFormula"></param>
		/// <param name="cellType"></param>
		/// <param name="rowStyleKey"></param>
		public void SetOneRowExcelCells(ISheet sheet, DataTable dataTable, int tableIndex, List<ExcelCellParam> param,
			ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null, Type cellType = null, string rowStyleKey = null)
		{
			var key = sheet.SheetName;

			SetOneRowExcelCells(sheet, key, dataTable, tableIndex, param, startColnum, rownum, rowStyle, isFormula, cellType, rowStyleKey);
		}

		/// <summary>
		/// 設定單排Excel
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="param"></param>
		/// <param name="startColnum"></param>
		/// <param name="rownum"></param>
		/// <param name="rowStyle"></param>
		/// <param name="isFormula"></param>
		public void SetOneRowExcelCells(ISheet sheet, List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1,
			Action<ICellStyle> rowStyle = null, bool? isFormula = null, Type cellType = null, string rowStyleKey = null)
		{
			var key = sheet.SheetName;

			SetOneRowExcelCells(sheet, key, new DataTable(), 0, param, startColnum, rownum, rowStyle, isFormula, cellType, rowStyleKey);
		}


		private void SetOneRowExcelCells(ISheet sheet, string groupKey, DataTable dataTable, int tableIndex,
			List<ExcelCellParam> param, ExcelColumns startColnum, int rownum = 1, Action<ICellStyle> rowStyle = null,
			bool? isFormula = null, Type rowCellType = null, string rowStyleKey = null)
		{
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var colnum = colIndex + startColnum;
				var col = param[colIndex];
				var isFormulaValue = col.IsFormula.HasValue ? col.IsFormula : isFormula;
				var styleKey = string.IsNullOrWhiteSpace(col.CellStyleKey) ? rowStyleKey : col.CellStyleKey;
				var cellType = col.CellValueType ?? rowCellType;
				SetExcelCell(sheet, groupKey, dataTable, tableIndex, col.ColumnName, colnum, rownum, col.CellValue,
					col.CellStyle, rowStyle, col.CellValueAction, isFormulaValue, cellType, styleKey);
			}
		}

		/// <summary>
		/// 設定多行Row
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="param"></param>
		/// <param name="startColnum"></param>
		/// <param name="startRownum"></param>
		/// <param name="rowStyle"></param>
		/// <param name="isFormula"></param>
		public void SetMultiRowsExcelCells(ISheet sheet, DataTable dataTable, List<ExcelCellParam> param,
			ExcelColumns startColnum, int startRownum = 1, Action<ICellStyle> rowStyle = null, bool? isFormula = null, Type rowCellType = null, string rowStyleKey = null)
		{
			startRownum = NormalizeRow(startRownum);
			var key = sheet.SheetName;

			for (int dtIndex = 0; dtIndex < dataTable.Rows.Count; dtIndex++)
			{
				var rownum = startRownum + dtIndex;
				SetOneRowExcelCells(sheet, key, dataTable, dtIndex, param, startColnum, rownum, rowStyle, isFormula, rowCellType, rowStyleKey);
			}
		}


		/// <summary>
		/// 設定表
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="tableCellParams"></param>
		/// <param name="startColnum"></param>
		/// <param name="rowNum"></param>
		/// <param name="headerStyle"></param>
		public void SetTableExcelCells(ISheet sheet, DataTable dataTable, List<TableCellParam> tableCellParams, ExcelColumns startColnum, int rowNum = 1, Action<ICellStyle> headerStyle = null, string headerRowStyleKey = null)
		{
			var headerParam = new List<ExcelCellParam>();
			var bodyParam = new List<ExcelCellParam>();
			sheet.CreateFreezePane((int)startColnum, rowNum);
			foreach (var p in tableCellParams)
			{
				headerParam.Add(new ExcelCellParam(p.HeaderName, colStyle: p.HeaderStyle, cellValueType: p.HeaderCellValueType));
				bodyParam.Add(new ExcelCellParam(p.CellValue, p.CellValueAction, p.CellStyle, p.IsFormula, p.CellValueType));
			}
			SetOneRowExcelCells(sheet, headerParam, startColnum, rowNum, headerStyle, rowStyleKey: headerRowStyleKey);
			SetMultiRowsExcelCells(sheet, dataTable, bodyParam, startColnum, rowNum + 1);
		}


		/// <summary>
		/// GetFormulaCellValue
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="colNum"></param>
		/// <param name="rowNum"></param>
		/// <returns></returns>
		public string GetFormulaCellValue(ISheet sheet, ExcelColumns colNum, int rowNum = 1)
		{
			if (sheet == null) throw new ArgumentNullException(nameof(sheet));
			
			rowNum = NormalizeRow(rowNum);
			rowNum = rowNum - 1;

			try
			{
				// 逐行讀取資料
				for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
				{
					IRow row = sheet.GetRow(rowIndex);
					if (row == null) continue;

					for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
					{
						ICell cell = row.GetCell(colIndex);
						if (rowIndex == rowNum && (int)colNum == colIndex)
						{
							IFormulaEvaluator evaluator = this.Workbook.GetCreationHelper().CreateFormulaEvaluator();
							DataFormatter formatter = new DataFormatter();
							var cellValue = formatter.FormatCellValue(cell, evaluator);
							return cellValue;
						}
					}
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"Error getting formula value: {ex.Message}");
				return null;
			}

			return null;
		}		public NpoiMemoryStream OutputExcelStream()
		{
			var ms = new NpoiMemoryStream();
			ms.AllowClose = false;
			Workbook.Write(ms);
			ms.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}

		public NpoiMemoryStream OutputExcelStream(Action<NpoiMemoryStream> modifyStream)
		{
			var ms = new NpoiMemoryStream();
			ms.AllowClose = false;
			Workbook.Write(ms);
			modifyStream(ms);
			ms.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}

		public void SetColumnWidth(ISheet sheet, ExcelColumns startCol, ExcelColumns endCol, double width)
		{
			for (int i = (int)startCol; i <= (int)endCol; i++)
			{
				sheet.SetColumnWidth(i, width * 256);
			}
		}


		public void RemovwRowRange(ISheet sheet, int startRow = 1, int endRow = 2)
		{
			startRow = NormalizeRow(startRow);
			endRow = NormalizeRow(endRow);
			
			if (endRow < startRow)
				endRow = startRow;
				
			startRow = startRow - 1;
			endRow = endRow - 1;
			
			for (int i = endRow; i >= startRow; i--)
			{
				IRow row = sheet.GetRow(i);
				if (row != null)
				{
					sheet.RemoveRow(row);
				}
			}
		}

		/// <summary>
		/// 插入圖片到指定位置
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="bytes"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		public void InsertPicture(ISheet sheet, byte[] bytes, ExcelColumns startCol, ExcelColumns endCol, int startRow, int endRow)
		{
			startRow = NormalizeRow(startRow);
			endRow = NormalizeRow(endRow);
			
			if (endRow < startRow)
				endRow = startRow;
				
			startRow = startRow - 1;
			endRow = endRow - 1;
			var picType = GetPictureType(bytes);

			int pictureIdx = Workbook.AddPicture(bytes, picType);

			// 建立繪圖patriarch
			IDrawing drawing = sheet.CreateDrawingPatriarch();

			// 設定圖片位置和大小 (起始行, 起始列, 結束行, 結束列)
			IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, (int)startCol, startRow, (int)endCol, endRow);

			// 插入圖片
			IPicture picture = drawing.CreatePicture(anchor, pictureIdx);

		}

		/// <summary>
		/// 根據圖片內容判斷格式
		/// </summary>
		private PictureType GetPictureType(byte[] imageData)
		{
			using (MemoryStream ms = new MemoryStream(imageData))
			using (System.Drawing.Image image = System.Drawing.Image.FromStream(ms))
			{
				if (image.RawFormat.Equals(ImageFormat.Jpeg))
					return PictureType.JPEG;
				else if (image.RawFormat.Equals(ImageFormat.Png))
					return PictureType.PNG;
				else if (image.RawFormat.Equals(ImageFormat.Gif))
					return PictureType.GIF;
				else if (image.RawFormat.Equals(ImageFormat.Bmp))
					return PictureType.BMP;
				else
					return PictureType.Unknown;
			}
		}

		/// <summary>
		/// Excel插入圖片到特定一格欄位
		/// </summary>
		/// <param name="wb">workbook</param>
		/// <param name="sheet">sheet</param>
		/// <param name="img">圖片</param>
		/// <param name="imgWidth">根據圖片調整的欄位寬度</param>
		/// <param name="col">插入行</param>
		/// <param name="row">插入列</param>
		/// <returns></returns>
		public IPicture SetPictureToExcellCell(ISheet sheet, byte[] img, int imgWidth, ExcelColumns col, int row = 1, AnchorType anchorType = AnchorType.DontMoveAndResize)
		{
			row = NormalizeRow(row);
			
			// **設定單元格大小**
			sheet.SetColumnWidth(col, imgWidth / 7); // 欄寬
			var picType = GetPictureType(img);
			
			// 插入圖片
			int pictureIdx = Workbook.AddPicture(img, picType);
			IDrawing drawing = sheet.CreateDrawingPatriarch();
			ICreationHelper helper = Workbook.GetCreationHelper();
			IClientAnchor anchor = helper.CreateClientAnchor();
			anchor.Col1 = (int)col; // 開始的欄
			anchor.Row1 = row - 1; // 開始的列
			anchor.AnchorType = anchorType;
			IPicture picture = drawing.CreatePicture(anchor, pictureIdx);
			return picture;
		}
	}

	public class NpoiMemoryStream : MemoryStream
	{
		public NpoiMemoryStream()
		{
			// We always want to close streams by default to
			// force the developer to make the conscious decision
			// to disable it.  Then, they're more apt to remember
			// to re-enable it.  The last thing you want is to
			// enable memory leaks by default.  ;-)
			AllowClose = true;
		}

		public bool AllowClose { get; set; }

		public override void Close()
		{
			if (AllowClose)
				base.Close();
		}
	}
}
