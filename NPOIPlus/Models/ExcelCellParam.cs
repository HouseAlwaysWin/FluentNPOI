using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public delegate T DefaultType<T>(T value);
	public delegate string CellValueActionType(ICell cell, object cellValue = null, ExcelColumns colnum = 0, int rownum = 0);

	public class ExcelCellParam
	{
		public readonly object CellValue;
		public readonly string ColumnName;
		public string CellStyleKey { get; set; }
		public readonly Action<ICellStyle> CellStyle;
		public readonly CellValueActionType CellValueAction;
		public readonly bool? IsFormula;
		public readonly Type CellValueType;

		public ExcelCellParam(object valueOrColName, CellValueActionType cellValueAction = null,
			Action<ICellStyle> colStyle = null, bool? isFormula = null, Type cellValueType = null, string cellStyleKey = null)
		{
			CellValueAction = cellValueAction;
			ColumnName = valueOrColName.ToString();
			CellValue = valueOrColName;
			CellStyle = colStyle;
			IsFormula = isFormula;
			CellValueType = cellValueType;
			CellStyleKey = cellStyleKey;
		}
	}
}
