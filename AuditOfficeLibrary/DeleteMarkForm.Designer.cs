namespace AuditOfficeLibrary
{
    partial class DeleteMarkForm
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
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.dgvMark = new System.Windows.Forms.DataGridView();
            this.colMark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMarkMean = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMark)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(124, 238);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 1;
            this.btnDelete.Text = "删除";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(205, 238);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // dgvMark
            // 
            this.dgvMark.AllowUserToAddRows = false;
            this.dgvMark.AllowUserToDeleteRows = false;
            this.dgvMark.AllowUserToResizeRows = false;
            this.dgvMark.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvMark.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMark.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colMark,
            this.colMarkMean});
            this.dgvMark.Location = new System.Drawing.Point(12, 12);
            this.dgvMark.Name = "dgvMark";
            this.dgvMark.ReadOnly = true;
            this.dgvMark.RowHeadersVisible = false;
            this.dgvMark.RowTemplate.Height = 23;
            this.dgvMark.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvMark.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvMark.Size = new System.Drawing.Size(268, 208);
            this.dgvMark.TabIndex = 2;
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
            // MFormSmartMark_Delete
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.dgvMark);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDelete);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MFormSmartMark_Delete";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "删除智能标志";
            this.Load += new System.EventHandler(this.MFormSmartMark_Delete_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMark)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridView dgvMark;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMark;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMarkMean;
    }
}