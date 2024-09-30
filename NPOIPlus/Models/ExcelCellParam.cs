using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public delegate T DefaultType<T>(T value);
	public delegate string CellValueActionType(ICell cell, object cellValue = null, int rownum = 0, ExcelColumns colnum = 0);

	public class ExcelCellParam
	{
		public object CellValue { get; set; }
		public string ColumnName { get; set; }
		public Action<ICellStyle> CellStyle { get; set; }
		public CellValueActionType CellValueAction { get; set; }
		public bool? IsFormula { get; set; }

		public ExcelCellParam(string columnName, object cellValue, CellValueActionType cellValueAction = null, Action<ICellStyle> style = null)
		{
			ColumnName = columnName;
			CellValueAction = cellValueAction;
			CellValue = cellValue;
			CellStyle = style;
		}
		public ExcelCellParam(string columnName, CellValueActionType cellValueAction = null, Action<ICellStyle> style = null, bool? isFormula = false)
		{
			CellValueAction = cellValueAction;
			ColumnName = columnName;
			CellStyle = style;
			IsFormula = isFormula;
		}
	}
}
