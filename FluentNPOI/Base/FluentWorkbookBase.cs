using NPOI.SS.UserModel;
using FluentNPOI.Helpers;
using System.IO;
using System.Collections.Generic;
using FluentNPOI.Models;

namespace FluentNPOI.Base
{
    public abstract class FluentWorkbookBase
    {
        protected IWorkbook _workbook;

        protected Dictionary<string, ICellStyle> _cellStylesCached;


        protected FluentWorkbookBase(IWorkbook workbook, Dictionary<string, ICellStyle> cellStylesCached)
        {
            _workbook = workbook;
            _cellStylesCached = cellStylesCached;
        }

        public IWorkbook GetWorkbook()
        {
            return _workbook;
        }

        protected void RegisterStyle(string name, System.Action<IWorkbook, ICellStyle> config)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            if (_cellStylesCached.ContainsKey(name)) return;

            var style = _workbook.CreateCellStyle();
            config?.Invoke(_workbook, style);
            _cellStylesCached[name] = style;
        }



        protected int NormalizeRow(int row)
        {
            // Convert 1-based row index (common for users) to 0-based, ensuring it's not negative
            if (row < 1) return 0;
            return row - 1;
        }

        public FluentMemoryStream ToStream()
        {
            var ms = new FluentMemoryStream();
            ms.AllowClose = false;
            _workbook.Write(ms);
            ms.Flush();
            ms.Seek(0, SeekOrigin.Begin);
            ms.AllowClose = true;
            return ms;
        }

        public IWorkbook SaveToPath(string filePath)
        {
            using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(outFile);
            }
            return _workbook;
        }
    }
}

