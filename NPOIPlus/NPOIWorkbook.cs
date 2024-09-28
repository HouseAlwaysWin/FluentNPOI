using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Data;

namespace NPOIPlus
{
	public delegate T DefaultType<T>(T value);

	public class NPOIWorkbook
	{
		public IWorkbook Workbook { get; set; }

		public DefaultType<int> SetDefaultIntCellValue = (value) => value;
		public DefaultType<double> SetDefaultDoubleCellValue = (value) => value;
		public DefaultType<bool> SetDefaultBoolCellValue = (value) => value;
		public DefaultType<string> SetDefaultStringCellValue = (value) => value;
		public DefaultType<DateTime> SetDefaultDateTimeCellValue = (value) => value;

		public Action<ICellStyle> SetGlobalCellStyle = (style) => { };
		public Action<ICellStyle> SetDefaultIntCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultDoubleCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultBoolCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultStringCellStyle = (value) => { };
		public Action<ICellStyle> SetDefaultDateTimeCellStyle = (value) => { };

		public NPOIWorkbook(IWorkbook workbook)
		{
			Workbook = workbook;
		}

		private void SetCellStyleBasedOnType(object value, ICellStyle style)
		{
			switch (value)
			{
				case int i:
					SetDefaultIntCellStyle?.Invoke(style);
					break;
				case double d:
					SetDefaultDoubleCellStyle?.Invoke(style);
					break;
				case bool b:
					SetDefaultBoolCellStyle?.Invoke(style);
					break;
				case string s:
					SetDefaultStringCellStyle?.Invoke(style);
					break;
				case DateTime dt:
					SetDefaultDateTimeCellStyle?.Invoke(style);
					break;
				default:
					SetDefaultStringCellStyle?.Invoke(style);
					break;
			}
		}


		private void SetCellValueBasedOnType(ICell cell, object value)
		{
			switch (value)
			{
				case int i:
					var intValue = SetDefaultIntCellValue(i);
					cell.SetCellValue(intValue);
					break;
				case double d:
					var doubleValue = SetDefaultDoubleCellValue(d);
					cell.SetCellValue(doubleValue);
					break;
				case bool b:
					var boolValue = SetDefaultBoolCellValue(b);
					cell.SetCellValue(boolValue);
					break;
				case string s:
					var stringValue = SetDefaultStringCellValue(s);
					cell.SetCellValue(stringValue);
					break;
				case DateTime dt:
					var dateValue = SetDefaultDateTimeCellValue(dt);
					cell.SetCellValue(dateValue);
					break;
				default:
					cell.SetCellValue(value?.ToString());
					break;
			}
		}

		/// <summary>
		/// For set single cell
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellValue"></param>
		/// <param name="colnum"></param>
		/// <param name="rownum"></param>
		/// <param name="param"></param>
		/// <exception cref="Exception"></exception>
		public void SetExcelCell(ISheet sheet, object cellValue, ExcelColumns colnum, int rownum, Action<ICellStyle> style = null)
		{
			if (rownum < 1) rownum = 1;
			IRow row = sheet.CreateRow(rownum - 1);
			ICell cell = row.CreateCell((int)colnum);
			SetCellStyle(cell, style);
			SetCellValueBasedOnType(cell, cellValue);
		}

		private void SetCellStyle(ICell cell, object cellValue, Action<ICellStyle> colStyle = null)
		{
			ICellStyle newCellStyle = Workbook.CreateCellStyle();
			SetGlobalCellStyle(newCellStyle);
			SetCellStyleBasedOnType(cellValue, newCellStyle);
			colStyle?.Invoke(newCellStyle);
			cell.CellStyle = newCellStyle;
		}

		/// <summary>
		/// For set single cell with datatable
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="dataTable"></param>
		/// <param name="tableIndex"></param>
		/// <param name="tableColName"></param>
		/// <param name="colnum"></param>
		/// <param name="rownum"></param>
		/// <param name="cellValue"></param>
		/// <exception cref="Exception"></exception>
		public void SetExcelCell(ISheet sheet, DataTable dataTable, int tableIndex, string tableColName, ExcelColumns colnum, int rownum = 1, object cellValue = null, Action<ICellStyle> style = null, bool isFormula = false)
		{
			if (rownum < 1) rownum = 1;
			IRow row = sheet.CreateRow(rownum - 1);
			ICell cell = row.CreateCell((int)colnum);
			var newValue = cellValue ?? dataTable.Rows[tableIndex][tableColName];
			SetCellStyle(cell, cellValue, style);
			if (isFormula)
			{
				cell.SetCellFormula(cellValue?.ToString());
				return;
			}
			SetCellValueBasedOnType(cell, newValue);
		}

		public void SetColExcelCells(ISheet sheet, DataTable dataTable, int tableIndex, List<ExcelCellParam> param, int rownum = 1, Action<ICellStyle> style = null, bool? isFormula = null)
		{
			for (int colIndex = 0; colIndex < param.Count; colIndex++)
			{
				var col = param[colIndex].Copy();
				if (isFormula != null) col.IsFormula = isFormula.Value;
				var newStyle = col.CellStyle ?? style;
				SetExcelCell(sheet, dataTable, tableIndex, col.ColumnName, (ExcelColumns)colIndex, rownum, col.CellValue, newStyle);
			}
		}

		public void SetRowExcelCells(ISheet sheet, DataTable dataTable, List<ExcelCellParam> param, int startRownum = 1, Action<ICellStyle> style = null, bool? isFormula = null)
		{
			if (startRownum < 1) startRownum = 1;

			for (int dtIndex = 0; dtIndex < dataTable.Rows.Count; dtIndex++)
			{
				var rownum = startRownum + dtIndex;
				SetColExcelCells(sheet, dataTable, dtIndex, param, rownum, style, isFormula);
			}
		}


	}
}
