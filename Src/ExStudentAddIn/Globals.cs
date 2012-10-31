using System;
using System.Collections.Generic;
using System.Text;
using IExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace ExStudentAddIn
{
    internal static class Globals
    {
        public static IExcelApplication ExcelApp
        {
            get;
            set;
        }
    }
}
