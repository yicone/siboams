namespace AuditOfficeLibrary
{
    partial class InsertMarkForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tpProject = new System.Windows.Forms.TabPage();
            this.dgvProjectMark = new System.Windows.Forms.DataGridView();
            this.colMark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMarkMean = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPageWorksheet = new System.Windows.Forms.TabPage();
            this.dgvWorksheetMark = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpReport = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMarkRpt = new System.Windows.Forms.TextBox();
            this.dgvReport = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpBal = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.txtMarkBal = new System.Windows.Forms.TextBox();
            this.dgvBal = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpOther = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMarkOther = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmboTypeForOther = new System.Windows.Forms.ComboBox();
            this.dgvOther = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpAnno = new System.Windows.Forms.TabPage();
            this.dgvAnno = new System.Windows.Forms.DataGridView();
            this.colAnoName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tpProject.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProjectMark)).BeginInit();
            this.tabPageWorksheet.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvWorksheetMark)).BeginInit();
            this.tpReport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReport)).BeginInit();
            this.tpBal.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBal)).BeginInit();
            this.tpOther.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOther)).BeginInit();
            this.tpAnno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAnno)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tpProject);
            this.tabControl1.Controls.Add(this.tabPageWorksheet);
            this.tabControl1.Controls.Add(this.tpReport);
            this.tabControl1.Controls.Add(this.tpBal);
            this.tabControl1.Controls.Add(this.tpOther);
            this.tabControl1.Controls.Add(this.tpAnno);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(284, 248);
            this.tabControl1.TabIndex = 0;
            // 
            // tpProject
            // 
            this.tpProject.Controls.Add(this.dgvProjectMark);
            this.tpProject.Location = new System.Drawing.Point(4, 21);
            this.tpProject.Name = "tpProject";
            this.tpProject.Padding = new System.Windows.Forms.Padding(3);
            this.tpProject.Size = new System.Drawing.Size(276, 223);
            this.tpProject.TabIndex = 0;
            this.tpProject.Text = "项目标志";
            this.tpProject.UseVisualStyleBackColor = true;
            // 
            // dgvProjectMark
            // 
            this.dgvProjectMark.AllowUserToAddRows = false;
            this.dgvProjectMark.AllowUserToDeleteRows = false;
            this.dgvProjectMark.AllowUserToResizeRows = false;
            this.dgvProjectMark.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvProjectMark.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvProjectMark.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colMark,
            this.colMarkMean,
            this.colSort});
            this.dgvProjectMark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvProjectMark.Location = new System.Drawing.Point(3, 3);
            this.dgvProjectMark.Name = "dgvProjectMark";
            this.dgvProjectMark.ReadOnly = true;
            this.dgvProjectMark.RowHeadersVisible = false;
            this.dgvProjectMark.RowTemplate.Height = 23;
            this.dgvProjectMark.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvProjectMark.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvProjectMark.Size = new System.Drawing.Size(270, 217);
            this.dgvProjectMark.TabIndex = 1;
            // 
            // colMark
            // 
            this.colMark.DataPropertyName = "Mark";
            this.colMark.HeaderText = "智能标志";
            this.colMark.Name = "colMark";
            this.colMark.ReadOnly = true;
            // 
            // colMarkMean
            // 
            this.colMarkMean.DataPropertyName = "MarkMean";
            this.colMarkMean.HeaderText = "含义";
            this.colMarkMean.Name = "colMarkMean";
            this.colMarkMean.ReadOnly = true;
            // 
            // colSort
            // 
            this.colSort.DataPropertyName = "Sort";
            this.colSort.HeaderText = "序号";
            this.colSort.Name = "colSort";
            this.colSort.ReadOnly = true;
            this.colSort.Visible = false;
            // 
            // tabPageWorksheet
            // 
            this.tabPageWorksheet.Controls.Add(this.dgvWorksheetMark);
            this.tabPageWorksheet.Location = new System.Drawing.Point(4, 21);
            this.tabPageWorksheet.Name = "tabPageWorksheet";
            this.tabPageWorksheet.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageWorksheet.Size = new System.Drawing.Size(276, 223);
            this.tabPageWorksheet.TabIndex = 1;
            this.tabPageWorksheet.Text = "底稿标志";
            this.tabPageWorksheet.UseVisualStyleBackColor = true;
            // 
            // dgvWorksheetMark
            // 
            this.dgvWorksheetMark.AllowUserToAddRows = false;
            this.dgvWorksheetMark.AllowUserToDeleteRows = false;
            this.dgvWorksheetMark.AllowUserToResizeRows = false;
            this.dgvWorksheetMark.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvWorksheetMark.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvWorksheetMark.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.dgvWorksheetMark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvWorksheetMark.Location = new System.Drawing.Point(3, 3);
            this.dgvWorksheetMark.Name = "dgvWorksheetMark";
            this.dgvWorksheetMark.ReadOnly = true;
            this.dgvWorksheetMark.RowHeadersVisible = false;
            this.dgvWorksheetMark.RowTemplate.Height = 23;
            this.dgvWorksheetMark.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvWorksheetMark.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvWorksheetMark.Size = new System.Drawing.Size(270, 217);
            this.dgvWorksheetMark.TabIndex = 2;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Mark";
            this.dataGridViewTextBoxColumn1.HeaderText = "智能标志";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "MarkMean";
            this.dataGridViewTextBoxColumn2.HeaderText = "含义";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "Sort";
            this.dataGridViewTextBoxColumn3.HeaderText = "序号";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.Visible = false;
            // 
            // tpReport
            // 
            this.tpReport.Controls.Add(this.label2);
            this.tpReport.Controls.Add(this.txtMarkRpt);
            this.tpReport.Controls.Add(this.dgvReport);
            this.tpReport.Location = new System.Drawing.Point(4, 21);
            this.tpReport.Name = "tpReport";
            this.tpReport.Padding = new System.Windows.Forms.Padding(3);
            this.tpReport.Size = new System.Drawing.Size(276, 223);
            this.tpReport.TabIndex = 2;
            this.tpReport.Text = "报表";
            this.tpReport.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "公式";
            // 
            // txtMarkRpt
            // 
            this.txtMarkRpt.Location = new System.Drawing.Point(41, 11);
            this.txtMarkRpt.Name = "txtMarkRpt";
            this.txtMarkRpt.Size = new System.Drawing.Size(229, 21);
            this.txtMarkRpt.TabIndex = 5;
            // 
            // dgvReport
            // 
            this.dgvReport.AllowUserToAddRows = false;
            this.dgvReport.AllowUserToDeleteRows = false;
            this.dgvReport.AllowUserToResizeRows = false;
            this.dgvReport.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvReport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvReport.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6});
            this.dgvReport.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dgvReport.Location = new System.Drawing.Point(3, 38);
            this.dgvReport.Name = "dgvReport";
            this.dgvReport.ReadOnly = true;
            this.dgvReport.RowHeadersVisible = false;
            this.dgvReport.RowTemplate.Height = 23;
            this.dgvReport.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvReport.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvReport.Size = new System.Drawing.Size(270, 182);
            this.dgvReport.TabIndex = 3;
            this.dgvReport.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvReport_RowEnter);
            this.dgvReport.SelectionChanged += new System.EventHandler(this.dgvReport_SelectionChanged);
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "Mark";
            this.dataGridViewTextBoxColumn4.HeaderText = "智能标志";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "MarkMean";
            this.dataGridViewTextBoxColumn5.HeaderText = "含义";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "Sort";
            this.dataGridViewTextBoxColumn6.HeaderText = "序号";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            this.dataGridViewTextBoxColumn6.Visible = false;
            // 
            // tpBal
            // 
            this.tpBal.Controls.Add(this.label1);
            this.tpBal.Controls.Add(this.txtMarkBal);
            this.tpBal.Controls.Add(this.dgvBal);
            this.tpBal.Location = new System.Drawing.Point(4, 21);
            this.tpBal.Name = "tpBal";
            this.tpBal.Padding = new System.Windows.Forms.Padding(3);
            this.tpBal.Size = new System.Drawing.Size(276, 223);
            this.tpBal.TabIndex = 3;
            this.tpBal.Text = "余额表";
            this.tpBal.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "公式";
            // 
            // txtMarkBal
            // 
            this.txtMarkBal.Location = new System.Drawing.Point(41, 6);
            this.txtMarkBal.Name = "txtMarkBal";
            this.txtMarkBal.Size = new System.Drawing.Size(229, 21);
            this.txtMarkBal.TabIndex = 4;
            // 
            // dgvBal
            // 
            this.dgvBal.AllowUserToAddRows = false;
            this.dgvBal.AllowUserToDeleteRows = false;
            this.dgvBal.AllowUserToResizeRows = false;
            this.dgvBal.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvBal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvBal.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9});
            this.dgvBal.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dgvBal.Location = new System.Drawing.Point(3, 33);
            this.dgvBal.Name = "dgvBal";
            this.dgvBal.ReadOnly = true;
            this.dgvBal.RowHeadersVisible = false;
            this.dgvBal.RowTemplate.Height = 23;
            this.dgvBal.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvBal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvBal.Size = new System.Drawing.Size(270, 187);
            this.dgvBal.TabIndex = 3;
            this.dgvBal.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBalance_RowEnter);
            this.dgvBal.SelectionChanged += new System.EventHandler(this.dgvBalance_SelectionChanged);
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "Mark";
            this.dataGridViewTextBoxColumn7.HeaderText = "智能标志";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "MarkMean";
            this.dataGridViewTextBoxColumn8.HeaderText = "含义";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "Sort";
            this.dataGridViewTextBoxColumn9.HeaderText = "序号";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.ReadOnly = true;
            this.dataGridViewTextBoxColumn9.Visible = false;
            // 
            // tpOther
            // 
            this.tpOther.Controls.Add(this.label4);
            this.tpOther.Controls.Add(this.txtMarkOther);
            this.tpOther.Controls.Add(this.label3);
            this.tpOther.Controls.Add(this.cmboTypeForOther);
            this.tpOther.Controls.Add(this.dgvOther);
            this.tpOther.Location = new System.Drawing.Point(4, 21);
            this.tpOther.Name = "tpOther";
            this.tpOther.Padding = new System.Windows.Forms.Padding(3);
            this.tpOther.Size = new System.Drawing.Size(276, 223);
            this.tpOther.TabIndex = 4;
            this.tpOther.Text = "其它";
            this.tpOther.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 197);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "标志";
            this.label4.Visible = false;
            // 
            // txtMarkOther
            // 
            this.txtMarkOther.Location = new System.Drawing.Point(41, 194);
            this.txtMarkOther.Name = "txtMarkOther";
            this.txtMarkOther.Size = new System.Drawing.Size(229, 21);
            this.txtMarkOther.TabIndex = 6;
            this.txtMarkOther.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "类别";
            // 
            // cmboTypeForOther
            // 
            this.cmboTypeForOther.FormattingEnabled = true;
            this.cmboTypeForOther.Items.AddRange(new object[] {
            "基本情况",
            "报表项目审定数",
            "纳税调增",
            "帐载金额",
            "准予税前扣除",
            "收入明细",
            "成本费用明细",
            "资产折旧摊销",
            "免税所得及减免税",
            "纳税调减",
            "其它"});
            this.cmboTypeForOther.Location = new System.Drawing.Point(41, 6);
            this.cmboTypeForOther.Name = "cmboTypeForOther";
            this.cmboTypeForOther.Size = new System.Drawing.Size(229, 20);
            this.cmboTypeForOther.TabIndex = 4;
            this.cmboTypeForOther.SelectedIndexChanged += new System.EventHandler(this.cmboTypeForOther_SelectedIndexChanged);
            // 
            // dgvOther
            // 
            this.dgvOther.AllowUserToAddRows = false;
            this.dgvOther.AllowUserToDeleteRows = false;
            this.dgvOther.AllowUserToResizeRows = false;
            this.dgvOther.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvOther.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvOther.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn12});
            this.dgvOther.Location = new System.Drawing.Point(3, 38);
            this.dgvOther.Name = "dgvOther";
            this.dgvOther.ReadOnly = true;
            this.dgvOther.RowHeadersVisible = false;
            this.dgvOther.RowTemplate.Height = 23;
            this.dgvOther.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvOther.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvOther.Size = new System.Drawing.Size(270, 150);
            this.dgvOther.TabIndex = 3;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "Mark";
            this.dataGridViewTextBoxColumn10.HeaderText = "智能标志";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.DataPropertyName = "Sort";
            this.dataGridViewTextBoxColumn12.HeaderText = "序号";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            this.dataGridViewTextBoxColumn12.ReadOnly = true;
            this.dataGridViewTextBoxColumn12.Visible = false;
            // 
            // tpAnno
            // 
            this.tpAnno.Controls.Add(this.dgvAnno);
            this.tpAnno.Location = new System.Drawing.Point(4, 21);
            this.tpAnno.Name = "tpAnno";
            this.tpAnno.Padding = new System.Windows.Forms.Padding(3);
            this.tpAnno.Size = new System.Drawing.Size(276, 223);
            this.tpAnno.TabIndex = 5;
            this.tpAnno.Text = "附注";
            this.tpAnno.UseVisualStyleBackColor = true;
            // 
            // dgvAnno
            // 
            this.dgvAnno.AllowUserToAddRows = false;
            this.dgvAnno.AllowUserToDeleteRows = false;
            this.dgvAnno.AllowUserToResizeRows = false;
            this.dgvAnno.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvAnno.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAnno.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colAnoName});
            this.dgvAnno.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvAnno.Location = new System.Drawing.Point(3, 3);
            this.dgvAnno.Name = "dgvAnno";
            this.dgvAnno.ReadOnly = true;
            this.dgvAnno.RowHeadersVisible = false;
            this.dgvAnno.RowTemplate.Height = 23;
            this.dgvAnno.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvAnno.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAnno.Size = new System.Drawing.Size(270, 217);
            this.dgvAnno.TabIndex = 3;
            // 
            // colAnoName
            // 
            this.colAnoName.DataPropertyName = "AnoName";
            this.colAnoName.HeaderText = "附注";
            this.colAnoName.Name = "colAnoName";
            this.colAnoName.ReadOnly = true;
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(139, 275);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(75, 23);
            this.btnInsert.TabIndex = 1;
            this.btnInsert.Text = "插入标签";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(221, 275);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // InsertMarkForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 312);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InsertMarkForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "插入智能标志";
            this.Load += new System.EventHandler(this.MFormSmartMark_Load);
            this.tabControl1.ResumeLayout(false);
            this.tpProject.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvProjectMark)).EndInit();
            this.tabPageWorksheet.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvWorksheetMark)).EndInit();
            this.tpReport.ResumeLayout(false);
            this.tpReport.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReport)).EndInit();
            this.tpBal.ResumeLayout(false);
            this.tpBal.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBal)).EndInit();
            this.tpOther.ResumeLayout(false);
            this.tpOther.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOther)).EndInit();
            this.tpAnno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAnno)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tpProject;
        private System.Windows.Forms.TabPage tabPageWorksheet;
        private System.Windows.Forms.DataGridView dgvProjectMark;
        private System.Windows.Forms.DataGridView dgvWorksheetMark;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TabPage tpReport;
        private System.Windows.Forms.TabPage tpBal;
        private System.Windows.Forms.TabPage tpOther;
        private System.Windows.Forms.DataGridView dgvReport;
        private System.Windows.Forms.DataGridView dgvBal;
        private System.Windows.Forms.DataGridView dgvOther;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMarkBal;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMarkRpt;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMark;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMarkMean;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSort;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMarkOther;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmboTypeForOther;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        private System.Windows.Forms.TabPage tpAnno;
        private System.Windows.Forms.DataGridView dgvAnno;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAnoName;
    }
}