using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Text;

namespace ExStudentAddIn
{
    public class HeadRowInfo
    {
        private int _rowNumber;
        private IDictionary<int, string> _headRowColumnMap;

        public HeadRowInfo(Worksheet sheet, int rowNumber, int columnCount)
        {
            _rowNumber = rowNumber;
            _headRowColumnMap = new Dictionary<int, string>();
            ParseHeadRow(sheet, rowNumber, columnCount);
        }

        private void ParseHeadRow(Worksheet sheet, int rowNumber, int columnCount)
        {
            Range headRow = sheet.Rows[rowNumber];
            for (int colIndex = 1; colIndex <= columnCount; colIndex++)
            {
                Range cell = headRow.Cells[colIndex];
                string colName = SheetHelper.GetCellText(cell);
                if (string.IsNullOrEmpty(colName) || string.IsNullOrEmpty(colName.Trim()))
                {
                    colName = SheetHelper.GetColumnName(colIndex);
                }
                _headRowColumnMap[colIndex] = colName;
            }
        }

        public string GetColumnName(int column)
        {
            string colName = null;
            if (!HeadColumns.TryGetValue(column, out colName))
            {
                colName = SheetHelper.GetColumnName(column);
            }
            return colName;
        }

        public IDictionary<int, string> HeadColumns
        {
            get
            {
                return _headRowColumnMap;
            }
        }

        public int RowNumber
        {
            get { return _rowNumber; }
            set { _rowNumber = value; }
        }
    }
}
