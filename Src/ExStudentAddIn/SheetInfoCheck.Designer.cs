namespace ExStudentAddIn
{
    partial class SheetInfoCheck
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblHeadRow = new System.Windows.Forms.Label();
            this.txtHeadRow = new System.Windows.Forms.TextBox();
            this.lblStudentIdColumn = new System.Windows.Forms.Label();
            this.lblApplySortColumn = new System.Windows.Forms.Label();
            this.lblStudentApplyListColumn = new System.Windows.Forms.Label();
            this.panelSheetInfo = new System.Windows.Forms.Panel();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.panelApplyProjectInfo = new System.Windows.Forms.Panel();
            this.lblApplyProjectInfoTip = new System.Windows.Forms.Label();
            this.panelResultInfo = new System.Windows.Forms.Panel();
            this.lblResultInfo = new System.Windows.Forms.Label();
            this.btnPrevious = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.cbProjectNameColumn = new System.Windows.Forms.ComboBox();
            this.cbStudentIdColumn = new System.Windows.Forms.ComboBox();
            this.cbApplySortColumn = new System.Windows.Forms.ComboBox();
            this.cbStudentApplyListColumn = new System.Windows.Forms.ComboBox();
            this.panelSheetInfo.SuspendLayout();
            this.panelApplyProjectInfo.SuspendLayout();
            this.panelResultInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblHeadRow
            // 
            this.lblHeadRow.AutoSize = true;
            this.lblHeadRow.Location = new System.Drawing.Point(21, 35);
            this.lblHeadRow.Name = "lblHeadRow";
            this.lblHeadRow.Size = new System.Drawing.Size(57, 17);
            this.lblHeadRow.TabIndex = 0;
            this.lblHeadRow.Text = "标题行：";
            // 
            // txtHeadRow
            // 
            this.txtHeadRow.Location = new System.Drawing.Point(229, 30);
            this.txtHeadRow.Name = "txtHeadRow";
            this.txtHeadRow.Size = new System.Drawing.Size(185, 22);
            this.txtHeadRow.TabIndex = 1;
            this.txtHeadRow.TextChanged += new System.EventHandler(this.txtHeadRow_TextChanged);
            // 
            // lblStudentIdColumn
            // 
            this.lblStudentIdColumn.AutoSize = true;
            this.lblStudentIdColumn.Location = new System.Drawing.Point(18, 136);
            this.lblStudentIdColumn.Name = "lblStudentIdColumn";
            this.lblStudentIdColumn.Size = new System.Drawing.Size(141, 17);
            this.lblStudentIdColumn.TabIndex = 2;
            this.lblStudentIdColumn.Text = "学生唯一标识所在列：";
            // 
            // lblApplySortColumn
            // 
            this.lblApplySortColumn.AutoSize = true;
            this.lblApplySortColumn.Location = new System.Drawing.Point(21, 191);
            this.lblApplySortColumn.Name = "lblApplySortColumn";
            this.lblApplySortColumn.Size = new System.Drawing.Size(155, 17);
            this.lblApplySortColumn.TabIndex = 4;
            this.lblApplySortColumn.Text = "申请的筛选标准所在列：";
            // 
            // lblStudentApplyListColumn
            // 
            this.lblStudentApplyListColumn.AutoSize = true;
            this.lblStudentApplyListColumn.Location = new System.Drawing.Point(24, 246);
            this.lblStudentApplyListColumn.Name = "lblStudentApplyListColumn";
            this.lblStudentApplyListColumn.Size = new System.Drawing.Size(169, 17);
            this.lblStudentApplyListColumn.TabIndex = 5;
            this.lblStudentApplyListColumn.Text = "学生申请项目列表所在列：";
            // 
            // panelSheetInfo
            // 
            this.panelSheetInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelSheetInfo.Controls.Add(this.cbStudentApplyListColumn);
            this.panelSheetInfo.Controls.Add(this.cbApplySortColumn);
            this.panelSheetInfo.Controls.Add(this.cbStudentIdColumn);
            this.panelSheetInfo.Controls.Add(this.cbProjectNameColumn);
            this.panelSheetInfo.Controls.Add(this.lblProjectName);
            this.panelSheetInfo.Controls.Add(this.lblApplySortColumn);
            this.panelSheetInfo.Controls.Add(this.lblHeadRow);
            this.panelSheetInfo.Controls.Add(this.txtHeadRow);
            this.panelSheetInfo.Controls.Add(this.lblStudentApplyListColumn);
            this.panelSheetInfo.Controls.Add(this.lblStudentIdColumn);
            this.panelSheetInfo.Location = new System.Drawing.Point(25, 13);
            this.panelSheetInfo.Name = "panelSheetInfo";
            this.panelSheetInfo.Size = new System.Drawing.Size(450, 341);
            this.panelSheetInfo.TabIndex = 7;
            // 
            // lblProjectName
            // 
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Location = new System.Drawing.Point(18, 81);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(141, 17);
            this.lblProjectName.TabIndex = 7;
            this.lblProjectName.Text = "申请项目名称所在列：";
            // 
            // panelApplyProjectInfo
            // 
            this.panelApplyProjectInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelApplyProjectInfo.Controls.Add(this.lblApplyProjectInfoTip);
            this.panelApplyProjectInfo.Location = new System.Drawing.Point(25, 13);
            this.panelApplyProjectInfo.Name = "panelApplyProjectInfo";
            this.panelApplyProjectInfo.Size = new System.Drawing.Size(450, 341);
            this.panelApplyProjectInfo.TabIndex = 8;
            // 
            // lblApplyProjectInfoTip
            // 
            this.lblApplyProjectInfoTip.AutoSize = true;
            this.lblApplyProjectInfoTip.Location = new System.Drawing.Point(22, 18);
            this.lblApplyProjectInfoTip.Name = "lblApplyProjectInfoTip";
            this.lblApplyProjectInfoTip.Size = new System.Drawing.Size(183, 17);
            this.lblApplyProjectInfoTip.TabIndex = 0;
            this.lblApplyProjectInfoTip.Text = "请输入各申请项目的名额数：";
            // 
            // panelResultInfo
            // 
            this.panelResultInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panelResultInfo.Controls.Add(this.lblResultInfo);
            this.panelResultInfo.Location = new System.Drawing.Point(25, 13);
            this.panelResultInfo.Name = "panelResultInfo";
            this.panelResultInfo.Size = new System.Drawing.Size(450, 341);
            this.panelResultInfo.TabIndex = 9;
            // 
            // lblResultInfo
            // 
            this.lblResultInfo.AutoSize = true;
            this.lblResultInfo.Location = new System.Drawing.Point(79, 84);
            this.lblResultInfo.Name = "lblResultInfo";
            this.lblResultInfo.Size = new System.Drawing.Size(43, 17);
            this.lblResultInfo.TabIndex = 0;
            this.lblResultInfo.Text = "结果：";
            // 
            // btnPrevious
            // 
            this.btnPrevious.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrevious.Location = new System.Drawing.Point(254, 435);
            this.btnPrevious.Name = "btnPrevious";
            this.btnPrevious.Size = new System.Drawing.Size(75, 25);
            this.btnPrevious.TabIndex = 10;
            this.btnPrevious.Text = "上一步";
            this.btnPrevious.UseVisualStyleBackColor = true;
            this.btnPrevious.Click += new System.EventHandler(this.btnPrevious_Click);
            // 
            // btnNext
            // 
            this.btnNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNext.Location = new System.Drawing.Point(335, 435);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 25);
            this.btnNext.TabIndex = 10;
            this.btnNext.Text = "下一步";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFinish.Location = new System.Drawing.Point(416, 435);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(75, 25);
            this.btnFinish.TabIndex = 10;
            this.btnFinish.Text = "完成";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Visible = false;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // cbProjectNameColumn
            // 
            this.cbProjectNameColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbProjectNameColumn.FormattingEnabled = true;
            this.cbProjectNameColumn.Location = new System.Drawing.Point(229, 78);
            this.cbProjectNameColumn.Name = "cbProjectNameColumn";
            this.cbProjectNameColumn.Size = new System.Drawing.Size(185, 24);
            this.cbProjectNameColumn.TabIndex = 8;
            // 
            // cbStudentIdColumn
            // 
            this.cbStudentIdColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudentIdColumn.FormattingEnabled = true;
            this.cbStudentIdColumn.Location = new System.Drawing.Point(229, 133);
            this.cbStudentIdColumn.Name = "cbStudentIdColumn";
            this.cbStudentIdColumn.Size = new System.Drawing.Size(185, 24);
            this.cbStudentIdColumn.TabIndex = 8;
            // 
            // cbApplySortColumn
            // 
            this.cbApplySortColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbApplySortColumn.FormattingEnabled = true;
            this.cbApplySortColumn.Location = new System.Drawing.Point(229, 188);
            this.cbApplySortColumn.Name = "cbApplySortColumn";
            this.cbApplySortColumn.Size = new System.Drawing.Size(185, 24);
            this.cbApplySortColumn.TabIndex = 8;
            // 
            // cbStudentApplyListColumn
            // 
            this.cbStudentApplyListColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudentApplyListColumn.FormattingEnabled = true;
            this.cbStudentApplyListColumn.Location = new System.Drawing.Point(229, 243);
            this.cbStudentApplyListColumn.Name = "cbStudentApplyListColumn";
            this.cbStudentApplyListColumn.Size = new System.Drawing.Size(185, 24);
            this.cbStudentApplyListColumn.TabIndex = 8;
            // 
            // SheetInfoCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(572, 512);
            this.Controls.Add(this.panelSheetInfo);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnPrevious);
            this.Controls.Add(this.panelApplyProjectInfo);
            this.Controls.Add(this.panelResultInfo);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SheetInfoCheck";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "SheetInfoCheck";
            this.panelSheetInfo.ResumeLayout(false);
            this.panelSheetInfo.PerformLayout();
            this.panelApplyProjectInfo.ResumeLayout(false);
            this.panelApplyProjectInfo.PerformLayout();
            this.panelResultInfo.ResumeLayout(false);
            this.panelResultInfo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblHeadRow;
        private System.Windows.Forms.TextBox txtHeadRow;
        private System.Windows.Forms.Label lblStudentIdColumn;
        private System.Windows.Forms.Label lblApplySortColumn;
        private System.Windows.Forms.Label lblStudentApplyListColumn;
        private System.Windows.Forms.Panel panelSheetInfo;
        private System.Windows.Forms.Panel panelApplyProjectInfo;
        private System.Windows.Forms.Label lblApplyProjectInfoTip;
        private System.Windows.Forms.Panel panelResultInfo;
        private System.Windows.Forms.Label lblResultInfo;
        private System.Windows.Forms.Button btnPrevious;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.ComboBox cbProjectNameColumn;
        private System.Windows.Forms.ComboBox cbStudentApplyListColumn;
        private System.Windows.Forms.ComboBox cbApplySortColumn;
        private System.Windows.Forms.ComboBox cbStudentIdColumn;
    }
}