using System;
using FluentNPOI;
using FluentNPOI.Models;
using FluentNPOI.Stages;

namespace FluentNPOIConsoleExample
{
    /// <summary>
    /// Read Examples - Reading Excel data
    /// </summary>
    internal partial class Program
    {
        #region Read Examples

        public static void ReadExcelExamples(FluentWorkbook fluent)
        {
            Console.WriteLine("\n========== 讀取 Excel 資料 ==========");

            // 讀取 BasicTableExample
            var sheet1 = fluent.UseSheet("BasicTableExample");
            Console.WriteLine("\n【BasicTableExample 標題行】:");
            for (ExcelCol col = ExcelCol.A; col <= ExcelCol.H; col++)
            {
                var headerValue = sheet1.GetCellValue<string>(col, 1);
                Console.Write($"{headerValue}\t");
            }
            Console.WriteLine();

            Console.WriteLine("\n【Sheet1 前3筆資料】:");
            for (int row = 2; row <= 4; row++)
            {
                var id = sheet1.GetCellValue<int>(ExcelCol.A, row);
                var name = sheet1.GetCellValue<string>(ExcelCol.B, row);
                var dateOfBirth = sheet1.GetCellValue<DateTime>(ExcelCol.C, row);
                Console.WriteLine($"Row {row}: ID={id}, Name={name}, Birth={dateOfBirth:yyyy-MM-dd}");
            }

            Console.WriteLine("\n========== 讀取完成 ==========\n");
        }

        #endregion
    }
}
