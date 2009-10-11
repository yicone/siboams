using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AuditOfficeLibrary;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using AuditPubLib;
using AuditPubLib.AuditDataSetSBMTableAdapters;
using AuditPubLib.AuditDataSetPROTableAdapters;
using System.Reflection;

namespace AuditOfficeLibrary
{
    public partial class InsertMarkForm : Form
    {
        //private OfficeWrap _officeWrap = null;
        private AuditDataSetSBM.SBM_OtherMarkDataTable _dtOtherMark = new AuditDataSetSBM.SBM_OtherMarkDataTable();
        private Excel.Application m_XlApp = null;
        private Word.Application m_WdApp = null;

        //public InsertMarkForm(OfficeWrap officeWrap)
        //{
        //    _officeWrap = officeWrap;

        //    InitializeComponent();
        //}

        public InsertMarkForm(Excel.Application xlApp)
        {
            m_XlApp = xlApp;
            InitializeComponent();
        }

        public InsertMarkForm(Word.Application wdApp)
        {
            m_WdApp = wdApp;
            InitializeComponent();
        }

        private void MFormSmartMark_Load(object sender, EventArgs e)
        {
            dgvProjectMark.AutoGenerateColumns = false;
            dgvWorksheetMark.AutoGenerateColumns = false;
            dgvReport.AutoGenerateColumns = false;
            dgvBal.AutoGenerateColumns = false;
            dgvOther.AutoGenerateColumns = false;
            dgvAnno.AutoGenerateColumns = false;

            AuditDataSetSBM.SBM_MarkDataTable dtMark = new AuditDataSetSBM.SBM_MarkDataTable();
            //AuditDataSetSBM.SBM_AnnoDataTable dtAnno = new AuditDataSetSBM.SBM_AnnoDataTable();
            DataTable dt = new DataTable();
            dt.Columns.Add("AnnoName");

            #region ȡ��������SBM_Mark���еı��
            try
            {
                new SBM_MarkTableAdapter().Fill(dtMark);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("���ط�'����'���͵ı��ʱ���ִ���:" + ex.Message);
            }

            DataView dv = new DataView(dtMark);
            dv.RowFilter = "Type = '��Ŀ'";
            dv.Sort = "Sort ASC";
            dgvProjectMark.DataSource = dv;

            dv = new DataView(dtMark);
            dv.RowFilter = "Type = '�׸�'";
            dv.Sort = "Sort ASC";
            dgvWorksheetMark.DataSource = dv;

            dv = new DataView(dtMark);
            dv.RowFilter = "Type = '����'";
            dv.Sort = "Sort ASC";
            dgvReport.DataSource = dv;

            dv = new DataView(dtMark);
            dv.RowFilter = "Type = '����' OR Type = '��������'";      //WXG 15:45 2007-3-8
            dv.Sort = "Sort,Type ASC";
            dgvBal.DataSource = dv; 
            #endregion

            #region ��ȡ������SBM_OtherMark��ı��
            try
            {
                new SBM_OtherMarkTableAdapter().Fill(_dtOtherMark);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("����'����'���͵ı��ʱ���ִ���:" + ex.Message);
            }

            dv = new DataView(_dtOtherMark);
            dv.Sort = "Sort ASC";
            dv.RowFilter = String.Format(@"Type = '����|{0}'", cmboTypeForOther.Text);
            dgvOther.DataSource = dv; 
            #endregion

            #region ��ȡ������SBM_Anno��ı��

            SBM_AnoTableAdapter daAno = new SBM_AnoTableAdapter();
            AuditDataSetSBM.SBM_AnoDataTable dtAno = null;
            try
            {
                dtAno = daAno.GetData();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("����'��ע'���͵ı��ʱ���ִ���:" + ex.Message);
            }

            dv = new DataView(dtAno);
            dv.Sort = "AnoName ASC";
            dgvAnno.DataSource = dv; 
            #endregion

            //������ͱ�������⴦��,�����ı�������ʾѡ���ĵ�Ԫ���еı�־,
            //��ʹ������ı�������ݺ�������ʱ,����Ԫ���еı�־ɾ��.
            try
            {
                if (m_XlApp != null)
                {
                    Excel.Range xlRange = m_XlApp.ActiveCell;
                    if (xlRange != null)
                    {
                        Excel.Name xlName = (Excel.Name)xlRange.Name;
                        if (xlName != null)
                        {
                            string name = xlName.Name;
                            string upperName = name.ToUpper();
                            if (upperName.StartsWith("BAL") || upperName.StartsWith("ABAL"))
                            {
                                txtMarkBal.Text = name;
                            }
                            else if (upperName.StartsWith("RPT"))
                            {
                                txtMarkRpt.Text = name;
                            }
                        }
                    }
                }

                //if (_officeWrap is ExcelWrap)
                //{
                //    ExcelWrap excelWrap = _officeWrap as ExcelWrap;
                //    if (excelWrap.ActiveCellName != null)
                //    {
                //        string cellName = excelWrap.ActiveCellName;
                //        string upperCellName = cellName.ToUpper();

                //        if (upperCellName.StartsWith("BAL") || upperCellName.StartsWith("ABAL"))    //WXG 15:45 2007-3-8
                //        {
                //            txtMarkBal.Text = cellName;
                //        }
                //        else if (upperCellName.StartsWith("RPT"))
                //        {
                //            txtMarkRpt.Text = cellName;
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace);
                //todo:��������ڷ���Name�������쳣.
                //throw ex;
            }

            #region business:����򿪵��ĵ����ڱ��������ı�־,����Ŀǰ�ĵ����Ѵ��ڵı�־�����ʾ���ݿ���ı�־
            /*
            try
            {
                AuditDataSetSBM.SBM_MarkRow drMark;
                SBM_MarkTableAdapter daMark = new SBM_MarkTableAdapter();
                if (_wordApp != null)
                {
                    AuditDataSetSBM.SBM_MarkDataTable dtMarkRPT = new AuditDataSetSBM.SBM_MarkDataTable();
                    AuditDataSetSBM.SBM_MarkDataTable dtMarkBAL = new AuditDataSetSBM.SBM_MarkDataTable();

                    foreach (Word.Bookmark bmk in _wordApp.ActiveDocument.Bookmarks)
                    {
                        bool hasClear = false;
                        string type;
                        if (daMark.GetDataByMark(bmk.Name).Count == 0)
                        {
                            if (bmk.Name.StartsWith("BAL"))
                            {
                                type = "����";
                            }
                            else if (bmk.Name.StartsWith("RPT"))
                            {
                                type = "����";
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            drMark = daMark.GetDataByMark(bmk.Name)[0];
                            type = drMark.Type;
                        }
                        
                        switch (type)
                        {
                            case "����":
                                if (!hasClear)
                                {
                                    dgvReport.DataSource = dtMarkRPT;
                                    hasClear = true;
                                }

                                AuditDataSetSBM.SBM_MarkRow drRPT = dtMarkRPT.NewSBM_MarkRow();
                                //drRPT.ItemArray = drMark.ItemArray;
                                drRPT.Mark = bmk.Name;
                                dtMarkRPT.AddSBM_MarkRow(drRPT);
                                break;
                            case "����":
                                if (!hasClear)
                                {
                                    hasClear = true;
                                    dgvBalance.DataSource = dtMarkBAL;
                                }

                                AuditDataSetSBM.SBM_MarkRow drBAL = dtMarkBAL.NewSBM_MarkRow();
                                //drBAL.ItemArray = drMark.ItemArray;
                                drBAL.Mark = bmk.Name;
                                dtMarkBAL.AddSBM_MarkRow(drBAL);
                                break;
                            default:
                                break;
                        }
                    }
                }
                else if (_excelApp != null)
                {
                    AuditDataSetSBM.SBM_MarkDataTable dtMarkRPT = new AuditDataSetSBM.SBM_MarkDataTable();
                    AuditDataSetSBM.SBM_MarkDataTable dtMarkBAL = new AuditDataSetSBM.SBM_MarkDataTable();

                    foreach (Excel.Name nme in _excelApp.ActiveWorkbook.Names)
                    {
                        bool hasClear = false;
                        string type;
                        if (daMark.GetDataByMark(nme.Name).Count != 1)
                        {
                            if (nme.Name.StartsWith("BAL"))
                            {
                                type = "����";
                            }
                            else if (nme.Name.StartsWith("RPT"))
                            {
                                type = "����";
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            drMark = daMark.GetDataByMark(nme.Name)[0];
                            type = drMark.Type;
                        }
                        switch (type)
                        {
                            case "����":
                                if (!hasClear)
                                {
                                    hasClear = true;
                                    dgvReport.DataSource = dtMarkRPT;
                                }

                                AuditDataSetSBM.SBM_MarkRow drRPT = dtMarkRPT.NewSBM_MarkRow();
                                //drRPT.ItemArray = drMark.ItemArray;
                                drRPT.Mark = nme.Name;
                                dtMarkRPT.AddSBM_MarkRow(drRPT);
                                break;
                            case "����":
                                if (!hasClear)
                                {
                                    hasClear = true;
                                    dgvBalance.DataSource = dtMarkBAL;
                                }

                                AuditDataSetSBM.SBM_MarkRow drBAL = dtMarkBAL.NewSBM_MarkRow();
                                //drBAL.ItemArray = drMark.ItemArray;
                                drBAL.Mark = nme.Name;
                                dtMarkBAL.AddSBM_MarkRow(drBAL);
                                break;
                            default:
                                break;
                        }
                    }
                }
                else
                {
                    throw new ArgumentOutOfRangeException("_wordApp and _excelApp are null");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("�������ܱ�־ʧ��:" + ex.Message);
            } */
            #endregion
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            string markName = null;
            bool validated = false;
            //ExcelWrap excelWrap = null;     //WXG 1.26
            DataGridViewRow dgvr = null;

            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    //��Ŀ
                    dgvr = dgvProjectMark.SelectedRows[0];
                    markName = dgvr != null ? dgvr.Cells[0].Value.ToString() : null;
                    validated = true;
                    break;
                case 1:
                    //�׸�
                    dgvr = dgvWorksheetMark.SelectedRows[0];
                    markName = dgvr != null ? dgvr.Cells[0].Value.ToString() : null;
                    validated = true;
                    break;
                case 2:
                    //����
                    markName = txtMarkRpt.Text.Trim();
                    //excelWrap = _officeWrap as ExcelWrap;
                    //if (excelWrap != null)
                    if(m_XlApp != null)
                    {
                        if (markName == "")
                        {
                            Excel.Name xlName = (Excel.Name)m_XlApp.ActiveCell.Name;
                            if (xlName != null)
                            {
                                string upperName = xlName.Name;
                                if (upperName.StartsWith("RPT"))
                                {
                                    xlName.Delete();
                                }
                            }
                        }
                        else
                        {
                            string upperName = markName.ToUpper();
                            string[] parts = markName.Split('_');
                            if (upperName.StartsWith("RPT"))
                            {
                                if (parts.Length >= 3)
                                {
                                    validated = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("��ʾ:�����־ֻ����Excel��ʹ��.");
                        return;
                    }
                    break;
                case 3:
                    //����+��������
                    markName = txtMarkBal.Text.Trim();
                    //excelWrap = _officeWrap as ExcelWrap;
                    //if (excelWrap != null)
                    //{
                    if(m_XlApp != null)
                    {
                        if (markName == "")
                        {
                            Excel.Name xlName = (Excel.Name)m_XlApp.ActiveCell.Name;
                            if (xlName != null)
                            {
                                string upperName = xlName.Name;
                                if (upperName.StartsWith("BAL") || upperName.StartsWith("ABAL"))
                                {
                                    xlName.Delete();
                                }
                            }
                        }
                        else  
                        {
                            string upperName = markName.ToUpper();
                            string[] parts = markName.Split('_');
                            if (upperName.StartsWith("BAL") || upperName.StartsWith("ABAL"))
                            {
                                if (parts.Length == 3)
                                {
                                    validated = true;
                                }
                                else if (parts.Length > 3)
                                {
                                    if (parts.Length == 5 && (parts[4] == "H" || parts[4] == "V") ||
                                        parts.Length == 7 && (parts[6] == "H" || parts[6] == "V") ||
                                        parts.Length == 6)
                                    {
                                        validated = true;
                                    }
                                }
                            }
                        }//WXG 15:45 2007-3-8
                    }
                    else
                    {
                        MessageBox.Show("��ʾ:�����־ֻ����Excel��ʹ��.");
                        return;
                    }
                    break;
                case 4:
                    //����
                    dgvr = dgvOther.SelectedRows[0];
                    markName = dgvr != null ? dgvr.Cells[0].Value.ToString() : null;
                    validated = true;
                    break;
                case 5:
                    //��ע
                    dgvr = dgvAnno.SelectedRows[0];
                    markName = dgvr != null ? dgvr.Cells[0].Value.ToString() : null;
                    validated = true;
                    break;
                default:
                    throw new ArgumentException();
            }


            if (validated)
            {
                //��ʼ����
                //_officeWrap.UpdateMark(markName);
                if (m_XlApp != null)
                {
                    m_XlApp.Names.Add(markName, m_XlApp.ActiveCell, 
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, 
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, 
                        Missing.Value);
                }
                else if (m_WdApp != null)
                {
                    object oRange = m_WdApp.Selection.Range;
                    m_WdApp.ActiveDocument.Bookmarks.Add(markName, ref oRange);
                }
            }
            else
            {
                MessageBox.Show("��־�ĸ�ʽ����ȷ.");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvBalance_SelectionChanged(object sender, EventArgs e)
        {
           
        }

        private void dgvReport_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void dgvBalance_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataRow dr = Common.GetSeletedSingleRow(dgvBal);
            if (dr != null)
            {
                txtMarkBal.Text = dr["Mark"].ToString();
            }
        }

        private void dgvReport_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataRow dr = Common.GetSeletedSingleRow(dgvReport);
            if (dr != null)
            {
                txtMarkRpt.Text = dr["Mark"].ToString();
            }
        }

        private void cmboTypeForOther_SelectedIndexChanged(object sender, EventArgs e)
        {
            //���������ʾ��־([����])
            DataView dv = new DataView(_dtOtherMark);
            dv.Sort = "Sort ASC";
            dv.RowFilter = String.Format(@"Type = '����|{0}'", cmboTypeForOther.Text);
            dgvOther.DataSource = dv;
        }
    }
}