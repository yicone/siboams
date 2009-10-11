using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AuditOfficeLibrary;
using AuditPubLib;
using AuditPubLib.AuditDataSetSBMTableAdapters;
using AuditPubLib.AuditDataSetPROTableAdapters;
using System.Reflection;


namespace AuditOfficeLibrary
{
    public partial class DeleteMarkForm : Form
    {
        private System.Data.DataTable m_DtMark = null;
        //private OfficeWrap m_OfficeWrap = null;
        private Excel.Application m_XlApp = null;
        private Word.Application m_WdApp = null;

        //public DeleteMarkForm(OfficeWrap officeWrap)
        //{
        //    m_OfficeWrap = officeWrap;

        //    InitializeComponent();

        //    //Initialize source DataTable
        //    m_DtMark = new System.Data.DataTable();
        //    m_DtMark.Columns.Add("Mark");            //HARDCODE
        //    m_DtMark.Columns.Add("MarkMean");
        //}

        public DeleteMarkForm(Excel.Application xlApp)
        {
            m_XlApp = xlApp;

            InitializeComponent();

            //Initialize source DataTable
            m_DtMark = new System.Data.DataTable();
            m_DtMark.Columns.Add("Mark");            //HARDCODE
            m_DtMark.Columns.Add("MarkMean");
        }

        public DeleteMarkForm(Word.Application wdApp)
        {
            m_WdApp = wdApp;

            InitializeComponent();

            //Initialize source DataTable
            m_DtMark = new System.Data.DataTable();
            m_DtMark.Columns.Add("Mark");            //HARDCODE
            m_DtMark.Columns.Add("MarkMean");
        }



        private void MFormSmartMark_Delete_Load(object sender, EventArgs e)
        {
            dgvMark.DataSource = m_DtMark;

            try
            {
                if (m_WdApp != null)
                {
                    foreach (Word.Bookmark bmk in m_WdApp.ActiveDocument.Bookmarks)
                    {
                        AddMarkRowToGrid(bmk.Name);
                    }
                }
                else if (m_XlApp != null)
                {
                    foreach (Excel.Name xlName in m_XlApp.ActiveWorkbook.Names)
                    {
                        AddMarkRowToGrid(xlName.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("�������ܱ�ǳ��ִ���:" + ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                #region ���ݷ���׼��
                SBM_MarkValueTableAdapter daMarkValue = new SBM_MarkValueTableAdapter();
                AuditDataSetSBM.SBM_MarkValueDataTable dtMarkValue;

                SBM_OtherMarkTableAdapter daOtherMark = new SBM_OtherMarkTableAdapter();
                AuditDataSetSBM.SBM_OtherMarkDataTable dtOtherMarkOneRow;
                AuditDataSetSBM.SBM_OtherMarkRow drOtherMark;

                SBM_AnoTableAdapter daAno = new SBM_AnoTableAdapter();
                AuditDataSetSBM.SBM_AnoDataTable dtAno;
                #endregion

                foreach (DataGridViewRow dgvr in dgvMark.SelectedRows)
                {
                    string markName = dgvr.Cells[0].Value.ToString();

                    //��������"Դ"���
                    if (markName[0] > 90 || markName[0] < 65)
                    {
                        if (markName.EndsWith("_"))
                        {
                            string actualMarkName = markName.Remove(markName.Length - 1);
                            //ɾ��SBM_MarkValue�����а����ñ�ǵļ�¼
                            dtMarkValue = daMarkValue.GetDataByMark(actualMarkName);
                            dtMarkValue.Clear();
                            daMarkValue.Update(dtMarkValue);
                            //��SBM_OtherMark������ñ��
                            dtOtherMarkOneRow = daOtherMark.GetDataByMark(actualMarkName);
                            if (dtOtherMarkOneRow.Count > 0)
                            {
                                drOtherMark = dtOtherMarkOneRow[0];
                                drOtherMark.Delete();
                                daOtherMark.Update(drOtherMark);
                            }
                        }
                    }

                    //ɾ����ע"Դ"���
                    if (markName.ToUpper().StartsWith("ANO"))
                    {
                        if (markName.EndsWith("_"))
                        {
                            string actualAnoName = markName.Remove(markName.Length - 1);
                            //TODO:�������ݿ���ɾ����־��ֵ (��������)

                            //��ɾ"Դ"���
                            dtAno = daAno.GetDataByKey(actualAnoName);
                            dtAno.Clear();
                            daAno.Update(dtAno);
                        }
                    }

                    //��Office�ĵ���ɾ�����(BookMark or Named Range)
                    //m_OfficeWrap.DeleteMark(markName);
                    if (m_XlApp != null)
                    {
                        Excel.Name xlName = m_XlApp.ActiveWorkbook.Names.Item(markName, Missing.Value, Missing.Value);
                        //if(xlName.RefersTo.ToString().Contains("REF!")) continue;

                        try
                        {
                            Excel.Range xlRange = xlName.RefersToRange;
                            //object oMark = mark;
                            //ɾ�����δ�滻ʱ��ʾ��ǵ�"<mark>"�ı�
                            string nmeText = xlRange.Text.ToString();

                            if (nmeText.Contains("<"))
                            {
                                int colonIndex = nmeText.IndexOf("<");

                                xlRange.Value2 = nmeText.Substring(0, colonIndex);
                            }
                        }
                        catch(Exception ex)
                        {
                            Log.Write(ex.Message + ex.StackTrace);
                        }
                        finally
                        {
                            xlName.Delete();
                        }
                    }
                    else if (m_WdApp != null)
                    {
                        object oMark = markName;
                        //ɾ�����δ�滻ʱ��ʾ��ǵ�"<mark>"�ı�
                        string bmkText = m_WdApp.ActiveDocument.Bookmarks.Item(ref oMark).Range.Text;

                        if (bmkText != null && bmkText.StartsWith("<") && bmkText.EndsWith(">"))
                        {
                            //ע��:����ı���ͬʱ,��ǩ�ѱ�ɾ��
                            //_wordSingleton.ActiveDocument.Bookmarks.get_Item(ref oMark).Range.Text = String.Empty;
                            //2000:
                            m_WdApp.ActiveDocument.Bookmarks.Item(ref oMark).Range.Text = String.Empty;
                        }
                        else
                        {
                            //_wordSingleton.ActiveDocument.Bookmarks.get_Item(ref oMark).Delete();
                            //2000:
                            m_WdApp.ActiveDocument.Bookmarks.Item(ref oMark).Delete();
                        }
                    }

                    //ˢ����ʾ
                    dgvMark.Rows.Remove(dgvr);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Method Helper
        private void AddMarkRowToGrid(string mark)
        {
            AuditDataSetSBM.SBM_MarkRow drMark;
            SBM_MarkTableAdapter daMark = new SBM_MarkTableAdapter();
            AuditDataSetSBM.SBM_MarkDataTable dtMarkOneRow;
            dtMarkOneRow = daMark.GetDataByMark(mark);

            if (dtMarkOneRow.Count > 0)
            {
                drMark = dtMarkOneRow[0];
                m_DtMark.Rows.Add(mark, drMark.MarkMean);
            }
            else
            {
                m_DtMark.Rows.Add(mark, "");
            }
        } 
        #endregion
    }
}