using System.Data;
using System;
using FluentNPOI;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using FluentNPOI.Models;
using System.Linq;
using FluentNPOI.Stages;
using FluentNPOI.Streaming.Mapping;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Table Write Examples - DataTable, Mapping, Summary
    /// </summary>
    internal partial class Program
    {
        #region Table Write Examples

        /// <summary>
        /// Example 1: Basic table - Demonstrates various field types using FluentMapping
        /// </summary>
        public static void CreateBasicTableExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 BasicTableExample (BasicTable)...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.ID).ToColumn(ExcelCol.A)
                .WithTitle("ID").WithTitleStyle("HeaderBlue").WithStyle("BodyGreen")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Name).ToColumn(ExcelCol.B)
                .WithTitle("名稱").WithTitleStyle("HeaderBlue").WithStyle("BodyGreen");
            mapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C)
                .WithTitle("生日").WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D)
                .WithTitle("是否活躍").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Boolean);
            mapping.Map(x => x.Score).ToColumn(ExcelCol.E)
                .WithTitle("分數").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Amount).ToColumn(ExcelCol.F)
                .WithTitle("金額").WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.Notes).ToColumn(ExcelCol.G)
                .WithTitle("備註").WithTitleStyle("HeaderBlue");

            mapping.Map(x => x.MaybeNull).ToColumn(ExcelCol.H)
                .WithTitle("可能為空").WithTitleStyle("HeaderBlue");

            fluent.UseSheet("BasicTableExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.H, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ BasicTableExample 建立完成");
        }

        /// <summary>
        /// Example 2: Summary table
        /// </summary>
        public static void CreateSummaryExample(FluentWorkbook fluent, List<ExampleData> testData)
        {
            Console.WriteLine("建立 Summary...");

            var mapping = new FluentMapping<ExampleData>();
            mapping.Map(x => x.Name).ToColumn(ExcelCol.A)
                .WithTitle("姓名11").WithTitleStyle("HeaderBlue");
            mapping.Map(x => x.Score).ToColumn(ExcelCol.B)
                .WithTitle("分數").WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency")
                .WithCellType(CellType.Numeric);
            mapping.Map(x => x.DateOfBirth).ToColumn(ExcelCol.C)
                .WithTitle("生日").WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map(x => x.IsActive).ToColumn(ExcelCol.D)
                .WithTitle("狀態").WithTitleStyle("HeaderBlue")
                .WithCellType(CellType.Boolean);
            mapping.Map(x => x.Notes).ToColumn(ExcelCol.E)
                .WithTitle("備註").WithTitleStyle("HeaderBlue")
            .WithStyle("HighlightYellow");

            fluent.UseSheet("Summary", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.E, 20)
                .SetTable(testData, mapping)
                .BuildRows();

            Console.WriteLine("  ✓ Summary 建立完成");
        }

        /// <summary>
        /// Example 3: Using DataTable as data source
        /// </summary>
        public static void CreateDataTableExample(FluentWorkbook fluent)
        {
            Console.WriteLine("建立 DataTableExample...");

            DataTable dataTable = new DataTable("StudentData");
            dataTable.Columns.Add("StudentID", typeof(int));
            dataTable.Columns.Add("StudentName", typeof(string));
            dataTable.Columns.Add("BirthDate", typeof(DateTime));
            dataTable.Columns.Add("IsEnrolled", typeof(bool));
            dataTable.Columns.Add("GPA", typeof(double));
            dataTable.Columns.Add("Tuition", typeof(decimal));
            dataTable.Columns.Add("Department", typeof(string));

            dataTable.Rows.Add(101, "張三", new DateTime(1998, 3, 15), true, 3.8, 25000m, "資訊工程");
            dataTable.Rows.Add(102, "李四", new DateTime(1999, 7, 20), true, 3.5, 22000m, "電機工程");
            dataTable.Rows.Add(103, "王五", new DateTime(1997, 11, 5), false, 2.9, 20000m, "機械工程");
            dataTable.Rows.Add(104, "趙六", new DateTime(2000, 1, 30), true, 3.9, 28000m, "資訊工程");
            dataTable.Rows.Add(105, "陳七", new DateTime(1998, 9, 12), true, 3.6, 23000m, "企業管理");
            dataTable.Rows.Add(106, "林八", new DateTime(1999, 5, 8), false, 3.2, 21000m, "財務金融");

            // 使用 DataTableMapping
            var mapping = new DataTableMapping();
            mapping.Map("StudentID").ToColumn(ExcelCol.A).WithTitle("學號")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Numeric);
            mapping.Map("StudentName").ToColumn(ExcelCol.B).WithTitle("姓名")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.String);
            mapping.Map("BirthDate").ToColumn(ExcelCol.C).WithTitle("出生日期")
                .WithTitleStyle("HeaderBlue").WithStyle("DateOfBirth");
            mapping.Map("IsEnrolled").ToColumn(ExcelCol.D).WithTitle("在學中")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Boolean);
            mapping.Map("GPA").ToColumn(ExcelCol.E).WithTitle("GPA")
                .WithTitleStyle("HeaderBlue").WithCellType(CellType.Numeric);
            mapping.Map("Tuition").ToColumn(ExcelCol.F).WithTitle("學費")
                .WithTitleStyle("HeaderBlue").WithStyle("AmountCurrency").WithCellType(CellType.Numeric);
            mapping.Map("Department").ToColumn(ExcelCol.G).WithTitle("科系")
                .WithTitleStyle("HeaderBlue")
                .WithValue((row, excelRow, col) => $"{row["StudentID"]}{excelRow}{col}{row["Department"]} hello")
                .WithStyle("BodyString")
                .WithCellType(CellType.String);

            fluent.UseSheet("DataTableExample", true)
                .SetColumnWidth(ExcelCol.A, ExcelCol.G, 20)
                .WriteDataTable(dataTable, mapping);

            Console.WriteLine("  ✓ DataTableExample 建立完成");
        }

        #endregion
    }
}
