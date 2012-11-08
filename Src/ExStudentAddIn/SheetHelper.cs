using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Text;

namespace ExStudentAddIn
{
    public static class SheetHelper
    {
        public static int GetColumnNumber(string colName)
        {
            string cleanColName = UpperAndTrim(colName);
            int minNumber = ((int)'A');
            int maxNumber = ((int)'Z');
            int baseNumber = maxNumber - minNumber + 1;
            int colNumber = 0;
            for (int i = 0; i < cleanColName.Length; i++)
            {
                char ch = cleanColName[i];
                colNumber += (((int)ch) - minNumber + 1) * (int)Math.Pow(baseNumber, cleanColName.Length - i - 1);
            }
            return colNumber;
        }

        public static string GetColumnName(int colNumber)
        {
            string colName = string.Empty;
            int minNumber = ((int)'A');
            int maxNumber = ((int)'Z');
            int baseNumber = maxNumber - minNumber + 1;
            int remainderNumber = colNumber;
            while (remainderNumber > 0)
            {
                int lowPositionNumber = remainderNumber % baseNumber;
                char numberChar = Convert.ToChar(lowPositionNumber - 1 + minNumber);
                colName = numberChar + colName;
                remainderNumber = remainderNumber / baseNumber;
            }
            return colName;
        }

        private static string UpperAndTrim(string colName)
        {
            if (!string.IsNullOrEmpty(colName))
            {
                return colName.Trim().ToUpper();
            }
            return string.Empty;
        }

        private static string TrimCellText(string cellText)
        {
            if (!string.IsNullOrEmpty(cellText))
            {
                return cellText.Trim();
            }
            return string.Empty;
        }

        public static string GetCellText(Range cell)
        {
            string text = null;
            if (cell != null)
            {
                text = cell.Text;
                if (text.StartsWith("#"))
                {
                    text = cell.get_Value();
                }
            }
            return TrimCellText(text);
        }
    }
}
