using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus
{
	public class ExcelCellParam
	{
		public object CellValue { get; set; }
		public string ColumnName { get; set; }
		public Action<ICellStyle> CellStyle { get; set; }
		public bool IsFormula { get; set; }


		public ExcelCellParam(object cellValue, string columnName = null, Action<ICellStyle> style = null, bool isFormula = false)
		{
			CellValue = cellValue;
			ColumnName = columnName;
			CellStyle = style;
			IsFormula = isFormula;
		}
	}
}
