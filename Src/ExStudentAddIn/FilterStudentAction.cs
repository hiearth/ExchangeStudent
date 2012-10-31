using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExStudentAddIn
{
    [Action(ControlId="MainGroupBtnGetResult")]
    internal class FilterStudentAction : Action
    {
        public override void DoAction()
        {
            try
            {
                Worksheet activeSheet = Globals.ExcelApp.ActiveSheet as Worksheet;
                int columnCount = activeSheet.UsedRange.Columns.Count;
                Range firstRow = activeSheet.Rows[1];
                for (int colIndex = 1; colIndex <= columnCount; colIndex++)
                {
                    Range cell = firstRow.Cells[colIndex];
                    string cellContent = cell.Text;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error occur");
            }
        }
    }
}
