using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AuditPubLib;
using AuditPubLib.AuditDataSetSBMTableAdapters;
using System.Reflection;

namespace AuditOfficeLibrary
{
    public partial class InsertOtherMarkForm : Form
    {
        AuditDataSetSBM.SBM_OtherMarkDataTable _dtOtherMark = new AuditDataSetSBM.SBM_OtherMarkDataTable();
        SBM_OtherMarkTableAdapter _daOtherMark = new SBM_OtherMarkTableAdapter();
        AuditDataSetSBM.SBM_MarkValueDataTable _dtMarkValue = new AuditDataSetSBM.SBM_MarkValueDataTable();
        SBM_MarkValueTableAdapter _daMarkValue = new SBM_MarkValueTableAdapter();
        //private int _projectId;
        //private OfficeWrap _officeWrap = null;
        private Excel.Application m_XlApp = null;
        private Word.Application m_WdApp = null;


        ////public MFormOtherMark_Edit(OfficeWrap officeWrap, int projectId, string mark, string markValue )
        //public InsertOtherMarkForm(OfficeWrap officeWrap, string mark)
        //{
        //    InitializeComponent();
            
        //    _officeWrap = officeWrap;
        //    txtMark.Text = mark;
        //    //txtMarkValue.Text = markValue;
        //}

        public InsertOtherMarkForm(Excel.Application xlApp, string mark)
        {
            InitializeComponent();

            m_XlApp = xlApp;
            txtMark.Text = mark;
            //txtMarkValue.Text = markValue;
        }

        public InsertOtherMarkForm(Word.Application wdApp, string mark)
        {
            InitializeComponent();

            m_WdApp = wdApp;
            txtMark.Text = mark;
            //txtMarkValue.Text = markValue;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //int markId = -1;
            if (cmboType.Text == "")
            {
                MessageBox.Show("����ѡ���ǵ����.");
                return;
            }

            if (txtMark.Text.Trim() == "")
            {
                MessageBox.Show("��ǲ���Ϊ��.");
                Common.ActiveTextBox(txtMark);
                return;
            }

            //����
            try
            {
                bool doubleInsert = true;
                AuditDataSetSBM.SBM_OtherMarkDataTable dtOtherMarkOneRow = _daOtherMark.GetDataByMark(txtMark.Text.Trim());
                if (dtOtherMarkOneRow.Count > 0)
                    doubleInsert = false;

                //�����ͬʱ����Ǹ���"_"�Ų��뵽�ĵ���
                //_officeWrap.UpdateMark(String.Format(@"{0}_", txtMark.Text.Trim()));
                string markName = String.Format(@"{0}_", txtMark.Text.Trim());
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
                    AuditDataSetSBM.SBM_OtherMarkRow drOtherMark = _dtOtherMark.NewSBM_OtherMarkRow();
                    drOtherMark.Mark = txtMark.Text.Trim();
                    drOtherMark.MarkMean = drOtherMark.Mark;
                    drOtherMark.Type = String.Format(@"����|{0}", cmboType.Text);
                    _dtOtherMark.AddSBM_OtherMarkRow(drOtherMark);
                    _daOtherMark.Update(_dtOtherMark);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ʱ���ִ���:" + ex.Message);
            }

            this.Close();
        }
    }
}