using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AuditPubLib;

namespace AuditOfficeLibrary
{
    public partial class InsertResultForm : Form
    {
        //private OfficeWrap m_officeWrap;
        private AuditPubLib.AuditDataSet.Ini_DictItemDataTable m_Dt;
        private Excel.Application m_XlApp = null;
        private Word.Application m_WdApp = null;

        //public InsertResultForm(OfficeWrap officeWrap)
        //{
        //    InitializeComponent();
        //    m_officeWrap = officeWrap;
        //}

        public InsertResultForm(Excel.Application xlApp)
        {
            m_XlApp = xlApp;
            InitializeComponent();
        }

        public InsertResultForm(Word.Application wdApp)
        {
            m_WdApp = wdApp;
            InitializeComponent();
        }

        private void InsertResultForm_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoGenerateColumns = false;
            m_Dt = DAL.GetAuditResultTable();
            cbType.SelectedIndex = 0;
            m_Dt.DefaultView.Sort = "Sort ASC";
            dataGridView1.DataSource = m_Dt.DefaultView;
        }

        private void cbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cbType.SelectedIndex == 0)
                m_Dt.DefaultView.RowFilter = "ParentDictID = '19'";
            else
                m_Dt.DefaultView.RowFilter = "ParentDictID = '20'";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(dataGridView1.SelectedRows.Count < 0 ) return;
            DataGridViewRow dgvr = dataGridView1.SelectedRows[0];
            //m_officeWrap.AppendText(dgvr.Cells[0].Value.ToString());
            string text = dgvr.Cells[0].Value.ToString();
            if (m_XlApp != null)
            {
                if (m_XlApp.ActiveCell != null)
                {
                    object o = m_XlApp.ActiveCell.Value2;
                    if (o != null)
                    {
                        string value2 = o.ToString();
                        if (value2.EndsWith(":") || value2.EndsWith("£º"))
                            text = value2 + text;
                    }
                }
                m_XlApp.ActiveCell.Value2 = text;
            }
            else if (m_WdApp != null)
            {
                m_WdApp.Selection.TypeText(text);
            }

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}