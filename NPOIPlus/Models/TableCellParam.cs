using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public class TableCellParam : ExcelCellParam
	{
		public readonly string HeaderName;
		public readonly Action<ICellStyle> HeaderStyle;
		public readonly Type HeaderCellValueType;
		public TableCellParam(string headerName, object bodyValue, CellValueActionType bodyCellValueActionType = null, Action<ICellStyle> bodyStyle = null, Action<ICellStyle> headerStyle = null, bool? isBodyFormula = null, Type bodyCellValueType = null, Type headerCellValueType = null) :
		base(bodyValue, bodyCellValueActionType, bodyStyle, isBodyFormula, bodyCellValueType)
		{
			HeaderName = headerName;
			HeaderStyle = headerStyle;
			HeaderCellValueType = headerCellValueType;
		}
	}
}
