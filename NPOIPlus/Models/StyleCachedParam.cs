using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIPlus.Models
{
	public class StyleCachedParam
	{
		public string Key { get; set; }
		public Action<ICellStyle> StyleAction { get; set; }
		public Type CellType { get; set; }

		public StyleCachedParam(string key, Action<ICellStyle> styleAction, Type cellType)
		{
			Key = key;
			StyleAction = styleAction;
			CellType = cellType;
		}
	}
}
