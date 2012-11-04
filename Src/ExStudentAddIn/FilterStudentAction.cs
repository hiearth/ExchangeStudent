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
                using (Form sheetInfoForm = new SheetInfoCheck())
                {
                    sheetInfoForm.StartPosition = FormStartPosition.CenterParent;
                    DialogResult dialogResult = sheetInfoForm.ShowDialog();
                    if (dialogResult == DialogResult.OK || dialogResult == DialogResult.Yes)
                    {

                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error occur");
            }
        }
    }
}
