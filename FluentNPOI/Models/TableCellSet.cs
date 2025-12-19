using NPOI.SS.UserModel;
using System;

namespace FluentNPOI.Models
{
    public class TableCellSet
    {
        public TableCellSet TitleCellSet { get; set; }
        public string CellName { get; set; }
        public object CellValue { get; set; }
        public Func<TableCellParams, object> SetValueAction { get; set; }
        public Func<TableCellParams, object> SetFormulaValueAction { get; set; }
        // Generic delegate (for ITable*Stage<T> use)
        public Delegate SetValueActionGeneric { get; set; }
        public Delegate SetFormulaValueActionGeneric { get; set; }
        public string CellStyleKey { get; set; }
        // Returns style configuration, StyleSetter executed only when needed
        public Func<TableCellStyleParams, CellStyleConfig> SetCellStyleAction { get; set; }
        public CellType CellType { get; set; }
    }
}

