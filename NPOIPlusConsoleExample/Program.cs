

using System.Data;
using System;
using System.IO;
using NPOIPlus;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIPlusConsoleExample
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// 打開 Excel 文件
			using (FileStream file = new FileStream(@"D:\Test.xlsx", FileMode.Open, FileAccess.Read))
			{
				// 1. 創建 DataTable
				DataTable dataTable = new DataTable("ExampleTable");

				// 2. 添加列 (列名稱與類型)
				dataTable.Columns.Add("ID", typeof(int));         // 整數類型的 ID 列
				dataTable.Columns.Add("Name", typeof(string));    // 字串類型的 Name 列
				dataTable.Columns.Add("DateOfBirth", typeof(DateTime)); // 日期類型的 DateOfBirth 列

				// 3. 添加數據行
				dataTable.Rows.Add(1, "Alice", new DateTime(1990, 1, 1));
				dataTable.Rows.Add(2, "Bob", new DateTime(1985, 5, 23));
				dataTable.Rows.Add(3, "Charlie", new DateTime(2000, 10, 15));

				NPOIWorkbook workbook = new NPOIWorkbook(new XSSFWorkbook(file));
				ISheet sheet1 = workbook.Workbook.GetSheet("Sheet1");

				workbook.SetExcelCell(sheet1, "你好", ExcelColumns.C, 5, (style) =>
				{
				});

			}
		}
	}

}