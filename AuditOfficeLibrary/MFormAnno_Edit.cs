using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AuditPubLib;
using AuditPubLib.AuditDataSetSBMTableAdapters;
using System.Diagnostics;

namespace AuditOfficeLibrary
{
    public partial class MFormAnno_Edit : Form
    {
        SBM_AnnoTableAdapter _daAnno = new SBM_AnnoTableAdapter();
        AuditDataSetSBM.SBM_AnnoDataTable _dtAnno = new AuditDataSetSBM.SBM_AnnoDataTable();
        private OfficeWrap _officeWrap = null;

        public MFormAnno_Edit(OfficeWrap officeWrap, string mark)
        {
            InitializeComponent();

            _officeWrap = officeWrap;
            txtAnno.Text = mark;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtAnno.Text.Trim() == "")
            {
                MessageBox.Show("附注不能为空.");
                Common.ActiveTextBox(txtAnno);
                return;
            }

            string annoName = txtAnno.Text.Trim();
            string directoryName = _officeWrap.DocWrap.DirectoryName;
            annoName = String.Format(@"Anno_{1}_{0}", annoName, directoryName);

            //判重
            AuditDataSetSBM.SBM_AnnoDataTable dtAnnoOneRow = _daAnno.GetDataByAnnoName(annoName);
            if (dtAnnoOneRow != null && dtAnnoOneRow.Count >= 1)
            {
                MessageBox.Show("该附注已存在.");
                return;
            }
            else
            {
                AuditDataSetSBM.SBM_AnnoRow drAnno = _dtAnno.NewSBM_AnnoRow();
                drAnno.AnnoName = annoName;
                drAnno.DirectoryName = directoryName;
                _dtAnno.AddSBM_AnnoRow(drAnno);

                try
                {
                    _daAnno.Update(drAnno);
                    _officeWrap.UpdateMark(String.Format(@"{0}_", annoName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("保存附注时出现错误:" + ex.Message);
                }
            }

            this.Close();
        }
    }
}