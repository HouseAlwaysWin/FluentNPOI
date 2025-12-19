namespace FluentNPOI.Models
{
    // Non-generic parameter type (compatible with old API)
    public class TableCellParams
    {
        public object CellValue { get; set; }
        public ExcelCol ColNum { get; set; }
        public int RowNum { get; set; }
        public object RowItem { get; set; }

        public T GetRowItem<T>()
        {
            return RowItem is T t ? t : default;
        }
    }

    public class TableCellParams<T>
    {
        public object CellValue { get; set; }
        public ExcelCol ColNum { get; set; }
        public int RowNum { get; set; }
        public T RowItem { get; set; }
    }
}

