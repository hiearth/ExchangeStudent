using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExStudentAddIn
{
    public partial class SheetInfoCheck : Form
    {
        public SheetInfoCheck()
        {
            InitializeComponent();
            Init();
            PrepareData();
        }

        private int _sheetColumnCount;
        private int _sheetRowCount;
        private ExchangeApply[] _exchangeApplies;
        private SheetInfo _sheetInfo;

        private void Init()
        {
            panelApplyProjectInfo.Visible = false;
            panelResultInfo.Visible = false;
            btnPrevious.Visible = false;
            btnFinish.Visible = false;
        }

        private void PrepareData()
        {
            Worksheet activeSheet = Globals.ExcelApp.ActiveSheet as Worksheet;
            _sheetInfo = new SheetInfo(activeSheet);

            txtHeadRow.Text = "1";
            IList<CellInfo> rowCells = _sheetInfo.GetRowCells(1);

            cbProjectNameColumn.DataSource = rowCells.ToArray();
            cbProjectNameColumn.DisplayMember = "Value";
            cbProjectNameColumn.ValueMember = "ColIndex";

            cbStudentIdColumn.DataSource = rowCells.ToArray();
            cbStudentIdColumn.DisplayMember = "Value";
            cbStudentIdColumn.ValueMember = "ColIndex";

            cbApplySortColumn.DataSource = rowCells.ToArray();
            cbApplySortColumn.DisplayMember = "Value";
            cbApplySortColumn.ValueMember = "ColIndex";

            cbStudentApplyListColumn.DataSource = rowCells.ToArray();
            cbStudentApplyListColumn.DisplayMember = "Value";
            cbStudentApplyListColumn.ValueMember = "ColIndex";
        }

        public ExchangeApply[] ExchangeApplies
        {
            get
            {
                if (_exchangeApplies == null)
                {
                    return new ExchangeApply[0];
                }
                return _exchangeApplies;
            }
            set { _exchangeApplies = value; }
        }

        public int SheetColumnCount
        {
            get
            {
                if (_sheetColumnCount <= 0)
                {
                    return int.MaxValue;
                }
                return _sheetColumnCount;
            }
            set
            {
                _sheetColumnCount = value;
            }
        }

        public int SheetRowCount
        {
            get
            {
                if (_sheetRowCount <= 0)
                {
                    return int.MaxValue;
                }
                return _sheetRowCount;
            }
            set
            {
                _sheetRowCount = value;
            }
        }

        private void DrawProjectsMaxCountUI(ApplyProject[] projects)
        {
            Point referLocation = new Point(lblApplyProjectInfoTip.Location.X, lblApplyProjectInfoTip.Location.Y + lblApplyProjectInfoTip.Height);
            for (int i = 0; i < projects.Length; i++)
            {
                // 生成label和textbox
                Label lbl = new Label();
                lbl.Text = projects[i].Name + ": ";
                lbl.Location = new Point(referLocation.X, referLocation.Y + 8);
                lbl.Width = 400;
                lbl.Height = 30;

                TextBox txtProject = new TextBox();
                txtProject.AccessibleName = "ProjectMaxCount";
                txtProject.Location = new Point(lbl.Location.X + lbl.Size.Width + 5, lbl.Location.Y);
                txtProject.Height = lbl.Height;
                txtProject.TextChanged += new EventHandler(txtProject_TextChanged);
                txtProject.Tag = projects[i];

                panelApplyProjectInfo.Controls.Add(lbl);
                panelApplyProjectInfo.Controls.Add(txtProject);

                referLocation = new Point(lbl.Location.X, lbl.Location.Y + lbl.Height);
            }
        }

        private void GenerateApplyProjectList()
        {
            List<ApplyProject> projects = new List<ApplyProject>();
            foreach (ExchangeApply applyItem in ExchangeApplies)
            {
                if (!projects.Contains(applyItem.Project))
                {
                    projects.Add(applyItem.Project);
                }
            }
            
        }

        private void txtProject_TextChanged(object sender, EventArgs e)
        {
            TextBox txtbox = sender as TextBox;
            if (txtbox != null)
            {
                if (!IntegerCheck(txtbox.Text, 0, int.MaxValue))
                {
                    ShowMessage("请输入大于等于 " + 0 + " 的整数！");
                }
            }
        }

        private void txtHeadRow_TextChanged(object sender, EventArgs e)
        {
            if (!IntegerCheck(txtHeadRow.Text, 1, SheetRowCount))
            {
                ShowMessage("请输入大于等于 " + 1 + ", 小于等于" + SheetRowCount + " 的整数！");
            }
        }

        private bool IntegerCheck(string value,int minValue, int maxValue)
        {
            int result = 0;
            if (int.TryParse(value, out result))
            {
                return result >= minValue && result <= maxValue;
            }
            return false;
        }

        private void ShowMessage(string msg)
        {
            MessageBox.Show(msg);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            try
            {
                if (panelSheetInfo.Visible)
                {
                    CaculateSheetInfo();

                    DrawProjectsMaxCountUI(_sheetInfo.Projects);

                    btnPrevious.Visible = true;
                    panelApplyProjectInfo.Visible = true;
                    panelSheetInfo.Visible = false;
                    panelResultInfo.Visible = false;
                }
                else if (panelApplyProjectInfo.Visible)
                {
                    CaculateApplyProjectInfo();

                    panelResultInfo.Visible = true;
                    panelSheetInfo.Visible = false;
                    panelApplyProjectInfo.Visible = false;
                }
                else if (panelResultInfo.Visible)
                {
                    CaculateResultInfo();
                    DrawResultInfo();

                    btnPrevious.Visible = false;
                    btnNext.Visible = false;
                    btnFinish.Visible = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void CaculateResultInfo()
        {
            _sheetInfo.GenerateFilterResult();
        }

        private void DrawResultInfo()
        {
            ExchangeApply[] applies = _sheetInfo.ExchangeApplies;
            Dictionary<int, string> outputHeadRow = GetOutPutHeadRow();
            Worksheet resultSheet = Globals.ExcelApp.Sheets.Add();
            resultSheet.Name = "Result";

            int headRowNumber = _sheetInfo.HeadRow.RowNumber;
            Range headRow = resultSheet.Rows[headRowNumber];
            for (int colIndex = 1; colIndex <= outputHeadRow.Count; colIndex++)
            {
                Range cell = headRow.Cells[colIndex];
                cell.Value = outputHeadRow[colIndex];
            }
            int contentRow = _sheetInfo.HeadRow.RowNumber + 1;
            foreach (ExchangeApply apply in applies)
            {
                Range row = resultSheet.Rows[contentRow];
                for (int colIndex = 1; colIndex <= outputHeadRow.Count; colIndex++)
                {
                    Range cell = row.Cells[colIndex];
                    string colName = outputHeadRow[colIndex];
                    string cellValue = string.Empty;
                    apply.ApplyInfo.TryGetValue(colName, out cellValue);
                    if (colName == "Pass")
                    {
                        cellValue = apply.Pass.ToString();
                    }
                    else if (colName == "Priority")
                    {
                        cellValue = apply.Priority.ToString();
                    }
                    cell.Value = cellValue;
                }
                contentRow++;
            }
        }

        private Dictionary<int, string> GetOutPutHeadRow()
        {
            Dictionary<int, string> outHeadRow = new Dictionary<int, string>();
            int passColIndex = 3;
            int priorityColIndex = 4;
            outHeadRow[passColIndex] = "Pass";
            outHeadRow[priorityColIndex] = "Priority";
            foreach (var headCol in _sheetInfo.HeadRow.HeadColumns)
            {
                int colIndex = headCol.Key;
                colIndex = colIndex >= passColIndex ? colIndex + 2 : colIndex;
                outHeadRow[colIndex] = headCol.Value;
            }
            return outHeadRow;
        }

        private void CaculateApplyProjectInfo()
        {
            foreach (Control ctl in panelApplyProjectInfo.Controls)
            {
                if (ctl.AccessibleName == "ProjectMaxCount")
                {
                    TextBox txtProjectMaxCount = ctl as TextBox;
                    if (txtProjectMaxCount != null)
                    {
                        ApplyProject project = txtProjectMaxCount.Tag as ApplyProject;
                        if (project != null)
                        {
                            project.MaxCount = int.Parse(txtProjectMaxCount.Text);
                        }
                    }
                }
            }
        }

        private void CaculateSheetInfo()
        {
            int headRow = int.Parse(txtHeadRow.Text);
            int projNameCol = int.Parse(cbProjectNameColumn.SelectedValue.ToString());
            int studentIdCol = int.Parse(cbStudentIdColumn.SelectedValue.ToString());
            int applySortCol = int.Parse(cbApplySortColumn.SelectedValue.ToString());
            int studentAppliesCol = int.Parse(cbStudentApplyListColumn.SelectedValue.ToString());
            _sheetInfo.ParseSheet(headRow, projNameCol, studentIdCol, applySortCol, studentAppliesCol);
        }

        private IList<CellInfo> GetHeadCells(int headRowNumber)
        {
            return _sheetInfo.GetRowCells(headRowNumber);
            //ComboBox cbNames = new ComboBox();
            //foreach(var cell in rowValues)
            //{
            //    CellInfo cellInfo = new CellInfo(cell.Key, cell.Value);
            //    cells.Add(cellInfo);
            //}
            //cbNames.ValueMember = "Index";
            //cbNames.DisplayMember = "Name";
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (panelApplyProjectInfo.Visible)
            {
                panelSheetInfo.Visible = true;
                panelApplyProjectInfo.Visible = false;
                panelResultInfo.Visible = false;
            }
            else if (panelResultInfo.Visible)
            {
                panelApplyProjectInfo.Visible = true;
                panelSheetInfo.Visible = false;
                panelResultInfo.Visible = false;
            }
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
