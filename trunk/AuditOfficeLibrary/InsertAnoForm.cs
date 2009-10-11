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
using System.Reflection;

namespace AuditOfficeLibrary
{
    public partial class InsertAnoForm : Form
    {
        SBM_AnoTableAdapter _daAno = new SBM_AnoTableAdapter();
        AuditDataSetSBM.SBM_AnoDataTable _dtAno = new AuditDataSetSBM.SBM_AnoDataTable();
        //private OfficeWrap _officeWrap = null;
        private Excel.Application m_XlApp = null;
        private Word.Application m_WdApp = null;

        //public InsertAnoForm(OfficeWrap officeWrap)
        //{
        //    InitializeComponent();

        //    _officeWrap = officeWrap;
        //}

        public InsertAnoForm(Excel.Application xlApp)
        {
            m_XlApp = xlApp;
            InitializeComponent();
        }

        public InsertAnoForm(Word.Application wdApp)
        {
            m_WdApp = wdApp;
            InitializeComponent();
        }


        private void MFormAno_Edit_Load(object sender, EventArgs e)
        {
            string directoryName = "";
            int id = Int32.Parse(System.IO.Path.GetFileNameWithoutExtension(m_XlApp.ActiveWorkbook.Name));
            AuditDataSetPRO.PRO_WorkSheetRow drWorksheet = DAL.FindWorksheet(id);
            if (drWorksheet != null)
            {
                int directoryId = drWorksheet.DirectoryID;
                directoryName = DAL.FindDirectory(directoryId).DirectoryName;
                
            }
            else
            {
                AuditDataSetTP.TP_WorkSheetRow drTemplate = DAL.FindTemplate(id);
                if (drTemplate != null)
                {
                    int directoryId = drTemplate.DirectoryID;
                    directoryName = DAL.FindTemplateDirectory(directoryId).DirectoryName;
                }
            }

            foreach (string ano in Ano.AnoCollection)
            {
                cbAno.Items.Add(String.Format(@"{0}_{1}", ano, directoryName));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region Input Validate
            string annoName = cbAno.Text.Trim();
            if (cbAno.Text.Trim() == "")
            {
                MessageBox.Show("附注不能为空.");
                return;
            }

            string[] array = annoName.Split('_');
            if (array.Length != 3)
            {
                MessageBox.Show("附注格式不正确.");
                return;
            }
            else
            {
                if (String.IsNullOrEmpty(array[2]))
                {
                    MessageBox.Show("附注格式不正确,没有输入报表项目.");
                }
            }
            #endregion

            //判重
            try
            {
                bool doubleInsert = true;
                AuditDataSetSBM.SBM_AnoDataTable dtAnoOneRow = _daAno.GetDataByKey(annoName);
                if (dtAnoOneRow.Count > 0)
                {
                    //MessageBox.Show("该附注已存在.");
                    //return;
                    doubleInsert = false;
                }

                //_officeWrap.UpdateMark(annoName + "_");
                string markName = annoName + "_";
                if (m_XlApp != null)
                {
                    m_XlApp.Names.Add(markName, m_XlApp.ActiveCell,
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value);
                }
                //else if (m_WdApp != null)
                //{ }

                if (doubleInsert)
                {
                    AuditDataSetSBM.SBM_AnoRow drAno = _dtAno.NewSBM_AnoRow();
                    drAno.AnoName = annoName;
                    _dtAno.AddSBM_AnoRow(drAno);
                    _daAno.Update(drAno);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存附注时出现错误:" + ex.Message);
            }

            this.Close();
        }
    }
}