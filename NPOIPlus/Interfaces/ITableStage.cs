using NPOI.SS.UserModel;
using NPOIPlus.Helpers;

namespace NPOIPlus
{
	public interface ITableStage<T>
	{
		FluentTableCellStage<T> BeginBodySet(string cellName);
		FluentTableHeaderStage<T> BeginTitleSet(string title);
		FluentTable<T> BuildRows();
	}
}

