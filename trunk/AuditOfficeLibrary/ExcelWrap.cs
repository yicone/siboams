using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using Excel;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AuditPubLib;
using Office;
using System.Data;
using Microsoft.VisualBasic;


namespace AuditOfficeLibrary
{
    public class ExcelWrap : OfficeWrap, IDisposable
    {
        private Excel.Application m_XlApp = null;
        private BackgroundWorker _saveWorker = new BackgroundWorker();
        private DocWrap _docWrap = null;
        private Form _excelFormThreadOwner = null;
        private bool _isAsyncSave = true;   //Ĭ��ִ���첽����
        private object _readonly = false;
        //private int _projectId;
        private bool _startQuit = false;

        public event DocumentBeforeSaveHandler DocumentBeforeSaveEvent;
        public event DocumentAfterSaveHandler DocumentAfterSaveEvent;

        //public override int ProjectId
        //{
        //    get { return _projectId; }
        //    set { _projectId = value; }
        //}

        public override DocWrap DocWrap
        {
            get { return _docWrap; }
        }

        public string ActiveCellName
        {
            get
            {
                if (m_XlApp.ActiveCell.Name != null)
                {
                    return ((Excel.Name)m_XlApp.ActiveCell.Name).Name;
                }

                return null;
            }
        }

        public Excel.Names Names
        {
            get
            {
                return m_XlApp.Names;
            }
        }


        public override List<Mark> Marks
        {
            get
            {
                List<Mark> marks = new List<Mark>();
                try
                {
                    Mark mark;
                    Excel.Range rng;
                    foreach (Excel.Name name in m_XlApp.Names)
                    {
                        try
                        {
                            rng = m_XlApp.get_Range(name.Name, Optional);
                            mark = new Mark(name.Name, rng.Row, rng.Column, rng.Worksheet.Index);
                            marks.Add(mark);
                        }
                        catch
                        {
                            //��ǲ�����,��������
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = "��ȡ�ĵ��ڵı�־���ִ���: " + ex.Message;
                    Debug.WriteLine(errMsg);
                    throw new Exception(errMsg, ex);
                }

                return marks;
            }
        }

        #region Obsolete! ԭ��ע��־
        /*public override Dictionary<string, string> AnnoDictionary
        {
            get
            {
                Dictionary<string, string> annoDictionary = new Dictionary<string, string>();
                foreach (Excel.Name nm in _excelSigleton.Names)
                {
                    if (nm.Name.ToUpper().StartsWith("ANNO"))
                    {
                        try
                        {
                            annoDictionary.Add(nm.Name, nm.RefersToRange.Text.ToString());
                        }
                        catch (Exception ex)
                        {
                            Debug.Assert(ex is COMException);
                            Debug.Assert(nm.RefersToRange != null);
                            //������Ч��Name��RefersToRange����ʱ���׳��쳣,
                            //�˴�����������"��ע��־"�ֵ�
                        }
                    }
                }
                return annoDictionary;
            }
        }*/

        #endregion

        public override Dictionary<string, string> RefedMarkDictionary
        {
            get
            {
                Dictionary<string, string> refedMarkDictionary = new Dictionary<string, string>();
                foreach (Excel.Name nm in m_XlApp.Names)
                {
                    //�����־������ĸ���Ǵ�д��Ӣ����ĸ,����Ϊ�ñ�־����"����"����
                    //todo:��Ҫʹ��"���ֿ�ͷ"��Ϊ"������־"������.
                    if (nm.Name[0] > 90 || nm.Name[0] < 65)
                    {
                        if (nm.Name.EndsWith("_"))
                        {
                            try
                            {
                                refedMarkDictionary.Add(nm.Name, nm.RefersToRange.Text.ToString());
                            }
                            catch (Exception ex)
                            {
                                Debug.Assert(ex is COMException);
                                Debug.Assert(nm.RefersToRange != null);
                                //������Ч��Name��RefersToRange����ʱ���׳��쳣,
                                //�˴�����������"������־"�ֵ�
                            }
                        }
                    }
                }

                return refedMarkDictionary;
            }
        }

        public bool IsAsyncSave
        {
            set { _isAsyncSave = value; }
        }

        public bool Readonly
        {
            get { return (bool)_readonly; }
            set { _readonly = value; }
        }

        public bool Visible
        {
            get { return m_XlApp.Visible; }
            set { m_XlApp.Visible = value; }
        }

        public override Dictionary<string, string> OtherMarkDictionary
        {
            get
            {
                Dictionary<string, string> otherMarkDictionary = new Dictionary<string, string>();
                foreach (Excel.Name nm in m_XlApp.Names)
                {
                    //�����־������ĸ���Ǵ�д��Ӣ����ĸ,����Ϊ�ñ�־����"����"����
                    if (nm.Name[0] > 90 || nm.Name[0] < 65)
                    {
                        otherMarkDictionary.Add(nm.Name, nm.RefersToRange.Text.ToString());
                    }
                }

                return otherMarkDictionary;
            }
        }

        /// <summary>
        /// �Զ���,��������DocumentAfterSave�¼�
        /// </summary>
        protected virtual void OnDocumentAfterSave(AfterSaveEventArgs e)
        {
            if (this.DocumentAfterSaveEvent != null)
            {
                DocumentAfterSaveEvent(this, e);
            }
        }

        /// <summary>
        /// �Զ���,��������DocumentBeforeSaveEvent�¼�
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnDocumentBeforeSave(BeforeSaveEventArgs e)
        {
            if (this.DocumentBeforeSaveEvent != null)
            {
                DocumentBeforeSaveEvent(this, e);
            }
        }

        #region ����

        public static ExcelWrap GetInstance(bool visable, Form excelFormThreadOwner)
        {
            return new ExcelWrap(visable, excelFormThreadOwner);
        }

        private ExcelWrap(bool visable, Form excelFormThreadOwner)
        {
            _excelFormThreadOwner = excelFormThreadOwner;
            _saveWorker.WorkerSupportsCancellation = true;
            _saveWorker.DoWork += new DoWorkEventHandler(_saveWorker_DoWork);
            _saveWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_saveWorker_RunWorkerCompleted);

            m_XlApp = new Excel.Application();
            m_XlApp.Visible = visable;

            #region ΪExcel Applicaton����ע���¼�������
            m_XlApp.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(_app_WorkbookBeforeSave);

            //�ر��ĵ�ǰ
            m_XlApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(_app_WorkbookBeforeClose);
            #endregion

            //��ӹ�����
            AddCommandBar();
        }

        void _saveWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //if (e.Error == null)
            //{
            //    if (!e.Cancelled)
            //    {
            //        MessageBox.Show("Ascnc operaion have completed successfully.");
            //    }
            //    else
            //    {
            //        MessageBox.Show("Ascnc operation was canceled because of background thread is busy.");
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Ascnc operation have completed, but that is error:" + e.Error);
            //}
        }

        /// <summary>
        /// ���CommandBar
        /// </summary>
        protected override void AddCommandBar()
        {
            try
            {
                //Ϊadd-in����һ��������
                _commandBarSBMMark = m_XlApp.CommandBars.Add("���ܱ�־",
                        MsoBarPosition.msoBarFloating,
                        Optional,
                        true);
                _commandBarSBMMark.Width = 100;

                base.AddCommandBarButtons();
            }
            catch (Exception ex)
            {
                string errMsg = "��ӹ��������ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #endregion

        #region Event Handle Method

        private int _WbCount = -1;
        /// <summary>
        /// �ڴ˴���Excel�ر�����
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Cancel"></param>
        void _app_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            bool isCloseApplication = true;
            if (_WbCount == -1)
            {
                _WbCount = Wb.Application.Workbooks.Count;
            }

            //��֤��Excel�򿪶��Workbookʱ,������Ϊ�ر����е�һ�����ر�����.
            if (_WbCount > 1)
            {
                isCloseApplication = false;
            }

            if (Path.GetExtension(Wb.Name) == ".txt")
            {
                Wb.Saved = true;    //��Ϊȫ�ļ���׼������ʱ�ı��ļ�������ִ���ٱ���
                return;
            }

            //�����Ƿ��޸Ĺ�,������"�Ƿ񱣴�"�ĶԻ���.
            //ShowMessageBoxCallback showMessageBoxCallback = new ShowMessageBoxCallback(MessageBox.Show);
            //DialogResult result = (DialogResult)_excelFormThreadOwner.Invoke(showMessageBoxCallback, new object[] { String.Format("�Ƿ񱣴��{0}�ĸ���?", Wb.Name), "Micorsoft Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation });
            if (!_startQuit)
            {
                DialogResult result;
                string text = String.Format("�Ƿ񱣴��{0}�ĸ���?", Wb.Name);
                string caption = "Microsoft Excel";
                MessageBoxButtons mbbs = MessageBoxButtons.YesNo;
                MessageBoxIcon mbi = MessageBoxIcon.Exclamation;
                
                IntPtr hwnd = Common.FindWindow("XLMAIN", null);
                if (hwnd != IntPtr.Zero)
                {
                    result = MessageBox.Show(new WindowWrap(hwnd), text, caption, mbbs, mbi);
                }
                else
                {
                    result = MessageBox.Show(text, caption, mbbs, mbi);
                }

                //DialogResult result;
                //Common.ShowMessageBox(String.Format("�Ƿ񱣴��{0}�ĸ���?", Wb.Name),
                //    "ȷ��",
                //    MessageBoxButtons.YesNo,
                //    _excelFormThreadOwner,
                //    out result);
                //Interaction.AppActivate("Microsoft Excel");

                switch (result)
                {
                    case DialogResult.Yes:
                        _isAsyncSave = false;
                        Wb.Save();
                        //wxg 2007��1��29�� 10:58
                        OnDocumentBeforeSave(new BeforeSaveEventArgs(_docWrap));
                        DoWorkAfterSaved(Wb);
                        if (isCloseApplication)
                        {
                            this.Dispose();
                        }
                        else
                        {
                            _WbCount--;
                        }
                        break;
                    case DialogResult.No:
                        Wb.Saved = true;
                        if (isCloseApplication)
                        {
                            this.Dispose();
                        }
                        else
                        {
                            _WbCount--;
                        }
                        break;
                    //case DialogResult.Cancel:
                    //    Cancel = true;
                    //    break;
                    default:
                        break;
                }
            }
        }

#if DEBUG
        private int m_ThreadNum = 0;
#endif
        void _app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            //todo:��Ҫ�ж�Wb�Ƿ����й��ĵ�
            if (_isAsyncSave && !SaveAsUI)
            {
                OnDocumentBeforeSave(new BeforeSaveEventArgs(_docWrap));

                try
                {
                    //�첽ִ�б���
                    if (_saveWorker.IsBusy)
                    {
                        _saveWorker.CancelAsync();
                        //try
                        //{
                        if (_saveWorker.IsBusy)
                        {
                            _saveWorker = new BackgroundWorker();
#if DEBUG
                            m_ThreadNum++;
#endif
                            _saveWorker.WorkerSupportsCancellation = true;
                            _saveWorker.DoWork += new DoWorkEventHandler(_saveWorker_DoWork);
                            _saveWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_saveWorker_RunWorkerCompleted);
                        }
                        //}
                        //catch (Exception ex)
                        //{
                        //    MessageBox.Show(ex.Message);
                        //}
                    }

                    _saveWorker.RunWorkerAsync(Wb);

                }
                catch (Exception ex)
                {
                    string errMsg = "BackroundWorker���ִ���:" + ex.Message;
                    //MessageBox.Show(errMsg);
                    Debug.WriteLine(errMsg);
                    throw new Exception(errMsg, ex);
                }
            }
        }

        /// <summary>
        /// �ڴ˼���DocumentAfterSave�¼�
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _saveWorker_DoWork(object sender, DoWorkEventArgs e)
        {
#if DEBUG
            if (String.IsNullOrEmpty(Thread.CurrentThread.Name))
            {
                Thread.CurrentThread.Name = "Background Worker Thread:" + m_ThreadNum;
            }
            else
            {
                MessageBox.Show(Thread.CurrentThread.Name);
            }
#endif

            Excel.Workbook wkb = e.Argument as Excel.Workbook;
            //Ǳ�ڵ�bug:����û��ڳ���򿪵�Excel���½������ĵ�����ֹxls��,
            //�򱣴���ĵ�ʱ���Ὣ�����½�/���ĵ�ʱ�����DocWrap�����Doc�����޸�Ϊ���ĵ���
            //�Ӷ���ʵ�ʱ��������ݿ���ֽڱ�͵����    
            DoWorkAfterSaved(wkb);
        }

        private void DoWorkAfterSaved(Excel.Workbook wkb)
        {
            if (Path.GetExtension(wkb.Name) == ".txt")
            {
                wkb.Saved = true;
                return;
            }

            try
            {
                int waitTime = 0;
                while (!wkb.Saved)
                {
                    Thread.Sleep(100);
                    waitTime+=100;
                    if (waitTime > 500)
                    {
                        return;
                    }

                    break;
                }

                OnDocumentAfterSave(new AfterSaveEventArgs(_docWrap));
            }
            catch (InvalidCastException)
            {
                //Debug.WriteLine("Save Workbook by SaveDialog starting..." + ex.Message);
                OnDocumentAfterSave(new AfterSaveEventArgs(_docWrap));
                return;
            }
            catch (COMException)
            {
                //Debug.WriteLine("Save Workbook by SaveDialog starting..." + ex.Message);
                OnDocumentAfterSave(new AfterSaveEventArgs(_docWrap));
                return;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }
        #endregion

        #region �ĵ�����
        /*
        public void OpenDoc(string filePathWithoutExtension, int id)
        {
            OpenDoc(filePathWithoutExtension);
            _docWrap.ID = id;
        }*/

        //TODO:NOTE:�Ժ��޸�Ϊ˽�л�ɾ��,Ŀǰ�ⲿ��Ϊ�׸����������׸�ʹ��(������)
        /// <summary>
        /// ���ĵ�
        /// </summary>
        /// <param name="filePathWithoutExtension"></param>
        /// <param name="isVisable"></param>
        public void OpenDoc(string filePathWithoutExtension)
        {
            _docWrap = new DocWrap();

            try
            {
                //xp: 15s
                //_docWrap.Doc = _excel.Workbooks.Open(filePathWithoutExtension,
                //                                Optional,
                //                                _readonly,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional,
                //                                Optional);

                //2000: 13s
                _docWrap.Doc = m_XlApp.Workbooks.Open(filePathWithoutExtension,
                                                    Optional,
                                                    _readonly,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional);

                _docWrap.EditState = 2; //Editing
                _docWrap.Path = filePathWithoutExtension;

                if (m_XlApp.Visible)
                {
                    //���
                    m_XlApp.WindowState = Excel.XlWindowState.xlMaximized;
                }

                Debug.WriteLine("�ĵ�����:���Excel Workbook��: " + m_XlApp.ActiveWorkbook.Name);
            }
            catch (Exception ex)
            {
                string errMsg = "��Workbook���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void OpenDoc(DocWrap docWrap)
        {
            _docWrap = docWrap;
            try
            {
                //2000: 13s
                _docWrap.Doc = m_XlApp.Workbooks.Open(DocWrap.Path,
                                                    Optional,
                                                    _readonly,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional,
                                                    Optional);

                docWrap.EditState = 2; //Editing

                if (m_XlApp.Visible)
                {
                    //���
                    m_XlApp.WindowState = Excel.XlWindowState.xlMaximized;
                }

                Debug.WriteLine("�ĵ�����:���Excel Workbook��: " + m_XlApp.ActiveWorkbook.Name);
            }
            catch (Exception ex)
            {
                string errMsg = "��Workbook���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        /*
        public void NewDoc(int docID, string filePathWithoutExtension, bool immediatelySave)
        {
            try
            {
                _docWrap.Doc = _excelSigleton.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet wks = (Excel.Worksheet)((Excel.Workbook)_docWrap.Doc).ActiveSheet;

                wks.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                wks.Name = "sheet1";

                _docWrap.ID = docID;
                _docWrap.EditState = 0;//New
                _docWrap.Path = filePathWithoutExtension;

                Thread.Sleep(1000);
                if (immediatelySave)
                {
                    SaveDocToLocal(filePathWithoutExtension);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "�½�Workbook���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }*/

        public void NewDoc(DocWrap docWrap, bool immediatelySave)
        {
            _docWrap = docWrap;
            try
            {
                _docWrap.Doc = m_XlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet wks = (Excel.Worksheet)((Excel.Workbook)_docWrap.Doc).ActiveSheet;

                wks.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                wks.Name = "sheet1";

                _docWrap.EditState = 0;//New

                Thread.Sleep(1000);
                if (immediatelySave)
                {
                    SaveDocToLocal(docWrap.Path);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "�½�Workbook���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        /// <summary>
        /// �����ĵ�(�����½�ʱʹ��)
        /// </summary>
        /// <param name="filePathWithoutExtension"></param>
        public void SaveDocToLocal(string filePathWithoutExtension)
        {
            try
            {
                //xp: 12s
                //((Excel.Workbook)_docWrap.Doc).SaveAs(filePathWithoutExtension,
                //        Optional,
                //        Optional,
                //        Optional,
                //        Optional,
                //        Optional,
                //        Excel.XlSaveAsAccessMode.xlExclusive,
                //        Optional,
                //        Optional,
                //        Optional,
                //        Optional,
                //        Optional);

                //2000:11s
                ((Excel.Workbook)_docWrap.Doc).SaveAs(filePathWithoutExtension,
                        Optional,
                        Optional,
                        Optional,
                        Optional,
                        Optional,
                        XlSaveAsAccessMode.xlNoChange,
                        Optional,
                        Optional,
                        Optional,
                        Optional);
            }
            catch (Exception ex)
            {
                string errMsg = "�½���Workbook���浽Temp�ļ���ʧ��: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #endregion

        #region CommandBarButton Event Handler
        protected override void btnInsertIndex_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //TODO:������ж��Ƿ��й��ĵ�,���ڷ��й��ĵ��в���ʱ,��鵽��Ӧ���й��ĵ���?
            if (_docWrap.Doc == null) return;

            try
            {
                object address = null;
                object textToDisplay = null;
                BeforeInsertIndexEventArgs e = new BeforeInsertIndexEventArgs(address, textToDisplay);
                //�������������¼�
                OnInsertIndex(e);

                if (e.Address != null && e.TextToDisplay != null)
                {
                    address = e.Address.ToString();
                    textToDisplay = e.TextToDisplay;
                    //������Word�Ĳ�����ͬ
                    ((Excel._Worksheet)(((Excel.Workbook)_docWrap.Doc).ActiveSheet)).Hyperlinks.Add(m_XlApp.ActiveCell, address.ToString(), Optional, Optional, textToDisplay);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "���뽻���������ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        protected override void btnInsertResult_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                InsertResultForm frm = new InsertResultForm(this);
                _excelFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _excelFormThreadOwner.Invoke(new MethodInvoker(_excelFormThreadOwner.SendToBack));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        protected override void btnDeleteMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                DeleteMarkForm frmDeleteMark = new DeleteMarkForm(this);

                _excelFormThreadOwner.Invoke(new MethodInvoker(frmDeleteMark.Show));
                _excelFormThreadOwner.Invoke(new MethodInvoker(_excelFormThreadOwner.SendToBack));
                //_excel.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected override void btnAddMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                InsertMarkForm frmInsertMark = new InsertMarkForm(this);

                _excelFormThreadOwner.Invoke(new MethodInvoker(frmInsertMark.Show));
                _excelFormThreadOwner.Invoke(new MethodInvoker(_excelFormThreadOwner.SendToBack));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected override void btnSaveMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                int rowIndex = m_XlApp.ActiveCell.Row;
                int colIndex = m_XlApp.ActiveCell.Column;
                string mark = "";
                if (colIndex - 1 >= 0)
                {
                    try
                    {
                        mark = ((Excel.Range)m_XlApp.Cells[rowIndex, colIndex - 1]).Text.ToString();
                    }
                    catch { }
                }

                //string markValue = _excelSigleton.ActiveCell.Text.ToString();

                InsertOtherMarkForm frm = new InsertOtherMarkForm(this, mark);
                _excelFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _excelFormThreadOwner.Invoke(new MethodInvoker(_excelFormThreadOwner.SendToBack));
            }
            catch (Exception ex)
            {
                string errMsg = "�����ǳ��ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        protected override void btnSaveAnno_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                InsertAnoForm frm = new InsertAnoForm(this);
                _excelFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _excelFormThreadOwner.Invoke(new MethodInvoker(_excelFormThreadOwner.SendToBack));
            }
            catch (Exception ex)
            {
                string errMsg = "���渽ע���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #endregion

        public override Byte[] GetDocBytes(string fileFullName)
        {
            string fullPath = fileFullName + ".xls";
            return base.GetDocBytes(fullPath);
        }

        //NOTE!���ڱ������ʱ����.
        public override string GetDocText(string filePath)
        {
            List<string> txtFileNameCollection = new List<string>();
            StringBuilder sb = new StringBuilder();

            try
            {
                //using (ExcelWrap excelWrap = new ExcelWrap(false, null))
                //{
                //    //NOTE!!ȷ��DoWorkAfterSaved������������!
                //    excelWrap.IsAsyncSave = false;

                //    string fileFullName = OfficeWrap.GetFullName(filePath);
                //    excelWrap.OpenDoc(fileFullName);
                //    Excel.Workbook xlBook = (Excel.Workbook)excelWrap.DocWrap.Doc;

                //    //��CSV��ʽ��ȡWorkbook�е��ı�
                //    string txtFileName;
                //    foreach (Excel.Worksheet xlSheet in xlBook.Worksheets)
                //    {
                //        txtFileName = Common.GetTempDirectoryPath() + Guid.NewGuid().ToString() + ".txt";

                //        //2000:9s
                //        xlSheet.SaveAs(txtFileName, XlFileFormat.xlCSV,
                //                Optional, Optional, Optional, Optional, Optional, Optional, Optional);

                //        txtFileNameCollection.Add(txtFileName);
                //    }
                //}

                string tempFilePath = Common.GetTempDirectoryPath() + Common.NewId();
                m_XlApp.ActiveWorkbook.SaveCopyAs(tempFilePath);
                using (ExcelWrap excelWrap = new ExcelWrap(false, null))
                {
                    //ȷ��DoWorkAfterSaved������������.
                    excelWrap.IsAsyncSave = false;
                    excelWrap.OpenDoc(tempFilePath);
                    Excel.Workbook xlBook = (Excel.Workbook)excelWrap.DocWrap.Doc;
                    string txtFileName;
                    foreach (Excel.Worksheet xlSheet in xlBook.Worksheets)
                    {
                        txtFileName = Common.GetTempDirectoryPath() + Common.NewId() + ".txt";
                        //��CSV��ʽ��ȡWorkbook�е��ı�
                        xlSheet.SaveAs(txtFileName, Excel.XlFileFormat.xlCSV,
                            Optional, Optional, Optional, Optional, Optional, Optional, Optional);

                        txtFileNameCollection.Add(txtFileName);
                    }
                }

                foreach (string fileName in txtFileNameCollection)
                {
                    using (FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        TextReader reader = new StreamReader(stream, Encoding.Default);//Encoding����������Ľ��
                        sb.Append(reader.ReadToEnd());
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "��ȡȫ�ı����ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return sb.ToString();
        }

        //NOTE!���ڱ������ʱ����.
        public override List<string> GetCrossRefList()
        {
            List<string> crossRefIdCollection = new List<string>();
            Excel.Workbook wb = (Excel.Workbook)_docWrap.Doc;

            if (!(Path.GetExtension(wb.Name) == ".txt"))//������������͵��ļ�,Excel����Ӧ��ʲô����?
            {
                try
                {
                    foreach (Excel.Hyperlink hl in ((Excel._Worksheet)wb.ActiveSheet).Hyperlinks)
                    {
                        string strRefedDocId = Path.GetFileNameWithoutExtension(hl.Address);
                        int refedDocId;
                        if (Int32.TryParse(strRefedDocId, out refedDocId))
                        {
                            //TODO:����Щ���ӵĵ�ַ�����õ������׸����ж�,�Ա��ִ˷�����ͨ����
                            //�������ظ��Ľ�������,�Է�ֹ���ĵ��򿪹������������õ��ļ�ʱ,
                            //�ظ�����ͬһ�ļ�,������IO�쳣;
                            //��������ĵ����������,����ͬ��;
                            if (!crossRefIdCollection.Contains(strRefedDocId) &&
                                strRefedDocId != Path.GetFileNameWithoutExtension(wb.Name))
                            {
                                crossRefIdCollection.Add(strRefedDocId);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = "��ȡ�����������ִ���: " + ex.Message;
                    Debug.WriteLine(errMsg);
                    throw new Exception(errMsg, ex);
                }
            }

            return crossRefIdCollection;
        }

        public override void AppendText(string text)
        {
            try
            {
                if (m_XlApp.ActiveCell != null)
                {
                    object o = m_XlApp.ActiveCell.Value2;
                    if (o != null)
                    {
                        string value2 = o.ToString();
                        if (value2.EndsWith(":") || value2.EndsWith("��"))
                            text = value2 + text;
                    }
                }
                m_XlApp.ActiveCell.Value2 = text;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void DeleteBALRPTMark()
        {
            try
            {
                Excel.Name nme = m_XlApp.ActiveCell.Name as Excel.Name;
                if (nme != null &&
                        (nme.Name.ToUpper().StartsWith("BAL") ||
                        nme.Name.ToUpper().StartsWith("RPT")))
                {
                    nme.Delete();
                }
            }
            catch
            {
                //����_excelSigleton.ActiveCell.Name ��Ϊnull,
                //�޷��ж�Cell��Range�Ƿ���Name,������������,ǿ��ɾ��.
            } 
        }

        /// <summary>
        /// ��Excel��ActiveCell�ϲ���NamedRange
        /// ����Ǳ���������־,�������ԭ����NamedRange���ٲ���,
        /// </summary>
        /// <param name="markName">�����String.Empty,��������</param>
        public override void UpdateMark(string markName)
        {
            try
            {
                if(markName != "")
                {
                    Excel.Name nme = m_XlApp.ActiveWorkbook.Names.Add(markName, m_XlApp.ActiveCell, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional);

                    nme.Visible = true;//����?
                }
            }
            catch (Exception ex)
            {
                string errMsg = "�����־���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mark"></param>
        public override void DeleteMark(string mark)
        {
            object oMark = mark;
            //ɾ�����δ�滻ʱ��ʾ��ǵ�"<mark>"�ı�
            string nmeText = m_XlApp.ActiveWorkbook.Names.Item(oMark, OfficeWrap.Optional, OfficeWrap.Optional).RefersToRange.Text.ToString();

            if (nmeText.Contains("<"))
            {
                int colonIndex = nmeText.IndexOf("<");

                m_XlApp.ActiveWorkbook.Names.Item(oMark, OfficeWrap.Optional, OfficeWrap.Optional).RefersToRange.Value2 = nmeText.Substring(0, colonIndex);
            }

            m_XlApp.ActiveWorkbook.Names.Item(oMark, OfficeWrap.Optional, OfficeWrap.Optional).Delete();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="marks"></param>
        public void ReplaceAllMarks(List<Mark> marks)
        {
            try
            {
                foreach (Mark mark in marks)
                {
                    //Excel.Range rng = (Excel.Range)_excelSigleton.Cells[mark.X, mark.Y];
                    Excel.Range rng = (Excel.Range)((Excel.Worksheet)m_XlApp.Sheets[mark.SheetIndex]).Cells[mark.X, mark.Y];
                    //Excel.Range rng;
                    //foreach (Excel.Name nme in _excelSigleton.Names)
                    //{
                    //    if (nme.Name != mark.Formula) continue;
                        
                    //    rng = nme.RefersToRange;
                        string value = rng.Value2 == null ? "" : rng.Value2.ToString();

                        int colonIndex = value.Contains(":") ? value.IndexOf(':') : value.IndexOf('��');
                        string iniValue = value.Substring(0, colonIndex + 1);

                        //��֤����ʱ���ַ���ʱ,Excel���Ὣ��Ԫ��ĸ�ʽ�Զ�����Ϊʱ������,���Ǳ����ı���ʽ
                        //rng.NumberFormatLocal = "��";
                        rng.Value2 = iniValue + mark.Value;
                    //    break;
                    //}
                }
            }
            catch (Exception ex)
            {
                string errMsg = "�滻��־ʱ���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #region ���������¼���
        public void ProcessTzflMark(int projectId, int accYear)
        {
            string strProjectId = projectId.ToString();
            string strYear = accYear.ToString();
            Excel.Range xlStartRange1 = null;
            int endRowIndex = -1;

            try
            {
                foreach (Excel.Name xlName in m_XlApp.Names)
                {
                    if (!xlName.Name.StartsWith("TZFL")) continue;

                    if (xlName.Name.EndsWith("END"))
                    {
                        endRowIndex = xlName.RefersToRange.Row;
                    }
                    else
                    {
                        xlStartRange1 = xlName.RefersToRange;
                    }

                    if (xlStartRange1 != null && endRowIndex != -1)
                    {
                        int startRowIndex = xlStartRange1.Row + 1;
                        int count = endRowIndex - startRowIndex;
                        int i2 = count;
                        while (i2 > 0)
                        {
                            xlStartRange1.get_Offset(i2, 0).EntireRow.Delete(Optional);
                            i2--;
                        }

                        //dtAdd cols:statement, value
                        List<System.Data.DataTable> adTableList = Pub_Function.GetAdjustData(strProjectId, strYear, xlName.Name);
                        Debug.Assert(adTableList.Count == 2);

                        Excel.Range xlStartInsertRowRange = xlStartRange1.get_Offset(1, 0);
                        int j2 = 0;
                        int count2 = adTableList[0].Rows.Count + adTableList[1].Rows.Count + 1;
                        if (adTableList[0].Rows.Count == 0 || adTableList[1].Rows.Count == 0)
                        {
                            count2 = count2 - 1;
                        }

                        while (j2 < count2)
                        {
                            xlStartInsertRowRange.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                            j2++;
                        }

                        Excel.Range xlStartRange2;
                        if (adTableList[0].Rows.Count > 0)
                        {
                            xlStartRange1.Value2 = "�ط����¼��";
                            xlStartRange2 = xlStartRange1.get_Offset(adTableList[0].Rows.Count + 1, 0);
                            Foo(adTableList[0], xlStartRange1);
                        }
                        else
                        {
                            xlStartRange2 = xlStartRange1;
                        }

                        if (adTableList[1].Rows.Count > 0)
                        {
                            xlStartRange2.Value2 = "������¼��";
                            Foo(adTableList[1], xlStartRange2);
                        }

                        break;
                    }
                }//end foreach
            }
            catch 
            {
                throw;
            }
        }

        private static void Foo(System.Data.DataTable dt, Excel.Range startRange)
        {
            Excel.Worksheet xlSheet = startRange.Worksheet;
            int i = 1;
            Excel.Range xlRange, xlValueRange;
            foreach (DataRow dr in dt.Rows)
            {
                if (dr[0].ToString().StartsWith("��"))  //borrow
                {
                    xlRange = (Excel.Range)xlSheet.get_Range(startRange.get_Offset(i, 0), startRange.get_Offset(i, 2));
                    xlRange.MergeCells = true;
                }
                else
                {
                    xlRange = (Excel.Range)xlSheet.get_Range(startRange.get_Offset(i, 0), startRange.get_Offset(i, 3));
                    xlRange.MergeCells = true;
                }

                //xlRange.ClearFormats();
                xlRange.Font.Bold = false; 
                xlRange.Value2 = dr[0].ToString();
                xlValueRange = xlRange.get_Offset(0, 1);
                xlValueRange.Font.Bold = false;
                xlValueRange.Font.Name = "Arial Narrow";
                xlValueRange.Font.Size = 9; 
                double value;
                if (double.TryParse(dr[1].ToString(), out value))
                {
                    xlValueRange.Value2 = value.ToString("N");
                }

                i++;
            }
        } 
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public System.Data.DataTable GetAnoValueDataTable()
        {
            #region Build Ano Value Table
            System.Data.DataTable dtAnoValue = new System.Data.DataTable();

            DataColumn colReportItem    = new DataColumn("ReportItem"); //������
            DataColumn colXM            = new DataColumn(Ano.AnoXM); //��
            DataColumn colNC            = new DataColumn(Ano.AnoNC, typeof(double));   //�����
            DataColumn colNM            = new DataColumn(Ano.AnoNM, typeof(double));   //��ĩ��
            DataColumn colJL            = new DataColumn(Ano.AnoJL, typeof(double));   //����
            DataColumn colDL            = new DataColumn(Ano.AnoDL, typeof(double));   //����
            DataColumn colTXT1          = new DataColumn(Ano.AnoTXT1);
            DataColumn colTXT2          = new DataColumn(Ano.AnoTXT2);
            DataColumn colTXT3          = new DataColumn(Ano.AnoTXT3);
            DataColumn colTXT4          = new DataColumn(Ano.AnoTXT4);
            DataColumn colVAL1          = new DataColumn(Ano.AnoVAL1, typeof(double));
            DataColumn colVAL2          = new DataColumn(Ano.AnoVAL2, typeof(double));

            dtAnoValue.Columns.AddRange(new DataColumn[] { colReportItem, colXM, colNC, colNM, colJL, colDL, 
                colVAL1, colVAL2, colTXT1, colTXT2, colTXT3, colTXT4});
            #endregion

            //List<string> anoDoubleList = new List<string>();
            //anoDoubleList.Add(Ano.AnoNC, Ano.AnoNM, Ano.AnoJL, Ano.AnoDL, Ano.AnoVAL1, Ano.AnoVAL2);
            //List<string> anoStringList = new List<string>();
            //anoStringList.Add(Ano.AnoTXT1, Ano.AnoTXT2, Ano.AnoTXT3, Ano.AnoTXT4);

            try
            {
                foreach (Excel.Name nmeXM in m_XlApp.Names)
                {
                    if (nmeXM.Name.StartsWith(Ano.AnoXM))
                    {
                        string[] array = nmeXM.Name.Split('_');
                        string reportItem = array[2];

                        if (reportItem == "") continue;

                        int i = 1;
                        string xmValue = "";
                        while (true)
                        {
                            Excel.Range rangeXM = null;
                            try
                            {
                                rangeXM = nmeXM.RefersToRange;
                            }
                            catch (Exception ex)
                            {
                                Debug.Assert(ex is COMException);
                                Debug.Assert(nmeXM.RefersToRange != null);
                                //������Ч��Name��RefersToRange����ʱ���׳��쳣,
                                //�˴�����������"��ע��־"�ֵ�
                            }
                            if (rangeXM.get_Offset(i, 0).Value2 != null)
                                xmValue = rangeXM.get_Offset(i, 0).Value2.ToString().Trim();
                            else
                                xmValue = "";

                            if (!String.IsNullOrEmpty(xmValue))
                            {
                                object oNcValue = GetAnoValueFromRange<double>(Ano.AnoNC, nmeXM, reportItem, i, 0d);
                                object oNmValue = GetAnoValueFromRange<double>(Ano.AnoNM, nmeXM, reportItem, i, 0d);
                                object oJlValue = GetAnoValueFromRange<double>(Ano.AnoJL, nmeXM, reportItem, i, 0d);
                                object oDlValue = GetAnoValueFromRange<double>(Ano.AnoDL, nmeXM, reportItem, i, 0d);
                                object oV1Value = GetAnoValueFromRange<double>(Ano.AnoVAL1, nmeXM, reportItem, i, 0d);
                                object oV2Value = GetAnoValueFromRange<double>(Ano.AnoVAL2, nmeXM, reportItem, i, 0d);
                                object oT1Value = GetAnoValueFromRange<string>(Ano.AnoTXT1, nmeXM, reportItem, i, String.Empty);
                                object oT2Value = GetAnoValueFromRange<string>(Ano.AnoTXT2, nmeXM, reportItem, i, String.Empty);
                                object oT3Value = GetAnoValueFromRange<string>(Ano.AnoTXT3, nmeXM, reportItem, i, String.Empty);
                                object oT4Value = GetAnoValueFromRange<string>(Ano.AnoTXT4, nmeXM, reportItem, i, String.Empty);
                                DataRow drWaitAdd = dtAnoValue.NewRow();

                                drWaitAdd[colReportItem] = reportItem;
                                drWaitAdd[colXM] = xmValue;
                                if (oNcValue != null)
                                    drWaitAdd[colNC] = (double)oNcValue;
                                if (oNmValue != null)
                                    drWaitAdd[colNM] = (double)oNmValue;
                                if (oJlValue != null)
                                    drWaitAdd[colJL] = (double)oJlValue;
                                if (oDlValue != null)
                                    drWaitAdd[colDL] = (double)oDlValue;
                                if (oV1Value != null)
                                    drWaitAdd[colVAL1] = (double)oV1Value;
                                if (oV2Value != null)
                                    drWaitAdd[colVAL2] = (double)oV2Value;
                                if (oT1Value != null)
                                    drWaitAdd[colTXT1] = oT1Value.ToString();
                                if (oT2Value != null)
                                    drWaitAdd[colTXT2] = oT2Value.ToString();
                                if (oT3Value != null)
                                    drWaitAdd[colTXT3] = oT3Value.ToString();
                                if (oT4Value != null)
                                    drWaitAdd[colTXT4] = oT4Value.ToString();

                                dtAnoValue.Rows.Add(drWaitAdd);
                            }//end if
                            else
                                break;

                            i++;
                        }
                    }//end if
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dtAnoValue;
        }


        private object GetAnoValueFromRange<T>(string ano, Excel.Name nmeXM, string reportItem, int i, T defaultValue)
        {
            //double anoValue = 0d;

            int columnOffset = 0;
            //object oAno = ano + "_";
            Excel.Name xlAnoName = null;

            try
            {
                foreach (Excel.Name xlName in m_XlApp.Names)
                {
                    if (xlName.Name.StartsWith(ano) && xlName.Name.EndsWith(reportItem + "_"))
                    {
                        xlAnoName = xlName;
                        break;
                    }
                }
            }
            catch { }

            //������ָ���ĸ�ע��־,��������
            if (xlAnoName == null)
            {
                return null;
            }
            else
            {
                columnOffset = xlAnoName.RefersToRange.Column - nmeXM.RefersToRange.Column;
                if (columnOffset != 0)
                {
                    Object o = nmeXM.RefersToRange.get_Offset(i, columnOffset).Value2;
                    
                    if (o != null)
                    {
                        try
                        {
                            T value = (T)o;
                            return value;
                        }
                        catch (InvalidCastException ex)
                        {

                        }
                    }
                }

                return defaultValue;
            }
        }

        //private String GetAnoStringVaue(string ano, Excel.Name nmeXM, string reportItem, int i)
        //{
        //    int columnOffset = 0;
        //    string anoValue = String.Empty;
        //    object oAno = ano + "_";
        //    Excel.Name nmeAno = null;
        //    try
        //    {
        //        foreach (Excel.Name nme in m_XlApp.Names)
        //        {
        //            if (nme.Name.StartsWith(ano) && nme.Name.EndsWith(reportItem + "_"))
        //            {
        //                nmeAno = nme;
        //                break;
        //            }
        //        }
        //    }
        //    catch { }

        //    if (nmeAno != null)
        //        columnOffset = nmeAno.RefersToRange.Column - nmeXM.RefersToRange.Column;
        //    if (columnOffset != 0)
        //    {
        //        Object o = nmeXM.RefersToRange.get_Offset(i, columnOffset).Value2;
        //        if (o != null)
        //            anoValue = nmeXM.RefersToRange.get_Offset(i, columnOffset).Value2.ToString();
        //    }

        //    return anoValue;
        //}

        //convert object to double without exception
        public static double ConvertToDouble(object o)
        {
            double value = 0d;
            if (o == null) return value;
            double.TryParse(o.ToString(), out value);
            return value;
        }

        /// <summary>
        /// �жϴ򿪵�Excelʵ�����Ƿ����ָ��������
        /// </summary>
        /// <param name="name">Named Range</param>
        /// <returns></returns>
        public static bool ExistName(Excel._Application excelApp, string name)
        {
            bool existName = false;
            try
            {
                Excel.Names names = ((Excel.Worksheet)excelApp.ActiveSheet).Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    if (names.Item(i, null, null).Name == name)
                    {
                        existName = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "���NamedRange�Ƿ����ʱ���ִ���:" + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return existName;
        }

        //����ճ��
        public void Paste()
        {
            try
            {
                Excel.Worksheet sheet = (Excel.Worksheet)m_XlApp.ActiveSheet;
                int i = 0;

                while (true)
                {
                    bool flag = false;
                    i++;
                    for (int j = 1; j < 6; j++)
                    {
                        if (((Excel.Range)sheet.Cells[i, j]).Text.ToString() != "")
                        {
                            flag = true;
                            break;
                        }
                    }

                    if (!flag)
                        break;
                }

                ((Excel.Worksheet)m_XlApp.ActiveSheet).Paste(m_XlApp.Cells[i, 1], null);
            }
            catch (Exception ex)
            {
                string errMsg = "ճ��ʱ���ִ���: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void CopyAndPasteDeclareSource(string preMarkName, ExcelWrap targetExcelWrap)
        {
            try
            {
                m_XlApp.DisplayAlerts = false;
                //Excel.Name nme = _excelSigleton.Names.Item(markName, Optional, Optional);
                //��Ч
                foreach (Excel.Name nme in m_XlApp.Names)
                {
                    if (nme.Name.StartsWith("DEC"))
                    {
                        nme.RefersToRange.Copy(Optional);
                        targetExcelWrap.PasteDeclareSource(nme.Name);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void PasteDeclareSource(string markName)
        {
            try
            {
                Excel.Name nme = m_XlApp.Names.Item(markName, Optional, Optional);
                //((Excel.Worksheet)_excelSigleton.ActiveSheet).Paste(nme.RefersToRange, Optional);
                //��ͬһ��Excel�и���ճ��
                //nme.RefersToRange.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Optional, Optional);

                //�Ӳ�ͬ��Excel�ܸ���ճ��
                //ActiveSheet.PasteSpecial Format:="�ı�", Link:=False, DisplayAsIcon:=False
                nme.RefersToRange.Worksheet.Activate();
                nme.RefersToRange.Select();
                ((Excel.Worksheet)nme.Application.ActiveSheet).PasteSpecial("�ı�", False, False, Optional, Optional, Optional);
            }
            catch (Exception)
            {

                throw;
            }

        }

        #region liyuan 2006-12-27 д���������
        //���ҵ�һ�в��ǵ�Ԫ�����,Ȼ����¶��п�ʼд������,���Ǻϲ���Ԫ�������,����
        public void WriteDiffData(System.Data.DataTable dt)
        {
            int iRowStart = 0;
            bool bFlag = false;
            double dbZCJF = 0;
            double dbZCDF = 0;
            double dbLRJF = 0;
            double dbLRDF = 0;
            string[] arrColumn = { "AdjustItemNo","WSIndexNum","AdjustResult","CodeName","RptNameID"};
            try
            {
                Excel.Worksheet sheet = (Excel.Worksheet)m_XlApp.ActiveSheet;
                while (iRowStart<20)
                {
                    iRowStart++;
                    if (((Excel.Range)sheet.Cells[iRowStart, 1]).Text.ToString().Trim() == "")
                    {
                        bFlag = true;
                        break;
                    }
                }
                if (bFlag)
                    iRowStart = iRowStart + 1;
                else
                    iRowStart = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < arrColumn.Length; j++)
                    {
                        if (j < 4)
                        {
                            ((Excel.Range)sheet.Cells[iRowStart + i, j + 1]).Value2 = dt.Rows[i][arrColumn[j]].ToString();
                        }
                        else
                        {
                            //1:�ʲ���ծ,2:����
                            if (dt.Rows[i][arrColumn[j]].ToString().Trim() == "1")
                            {
                                //�ʲ���ծ��
                                ((Excel.Range)sheet.Cells[iRowStart + i, j + 1]).Value2 = dt.Rows[i]["Debit"].ToString();
                                if (dt.Rows[i]["Debit"].ToString() != "")
                                {
                                    dbZCJF += double.Parse(dt.Rows[i]["Debit"].ToString());
                                }
                                ((Excel.Range)sheet.Cells[iRowStart + i, j + 2]).Value2 = dt.Rows[i]["Credit"].ToString();
                                if (dt.Rows[i]["Credit"].ToString() != "")
                                {
                                    dbZCDF += double.Parse(dt.Rows[i]["Credit"].ToString());
                                }                   
                            }
                            else
                            {
                                //�����
                                ((Excel.Range)sheet.Cells[iRowStart + i, j + 3]).Value2 = dt.Rows[i]["Debit"].ToString();
                                if (dt.Rows[i]["Debit"].ToString() != "")
                                {
                                    dbLRJF += double.Parse(dt.Rows[i]["Debit"].ToString());
                                }
                                ((Excel.Range)sheet.Cells[iRowStart + i, j + 4]).Value2 = dt.Rows[i]["Credit"].ToString();
                                if (dt.Rows[i]["Credit"].ToString() != "")
                                {
                                    dbLRDF += double.Parse(dt.Rows[i]["Credit"].ToString());
                                }
                            }
                        }
                    }
                }
                //���������Ϻ�,�Է�¼��,����ԭ�������úϲ���Ԫ��
                string strItemNo = dt.Rows[0]["AdjustItemNo"].ToString().Trim();
                int iSRowNO = iRowStart;
                Excel.Range rg = null;
                for (int i = iRowStart; i < iRowStart + dt.Rows.Count; i++)
                {
                    if (strItemNo != dt.Rows[i - iRowStart]["AdjustItemNo"].ToString())
                    {
                        if (i - 1 - iSRowNO > 0)
                        {
                            rg = sheet.get_Range(sheet.Cells[iSRowNO, 1], sheet.Cells[i - 1, 1]);
                            rg.MergeCells = true;
                            rg.WrapText = true;
                            rg = sheet.get_Range(sheet.Cells[iSRowNO, 3], sheet.Cells[i - 1, 3]);
                            rg.MergeCells = true;
                            rg.WrapText = true;
                            iSRowNO = i;
                        }
                    }
                    else
                    {
                        if (i != iSRowNO)
                        {
                            ((Excel.Range)sheet.Cells[i, 1]).Value2 = "";
                            ((Excel.Range)sheet.Cells[i, 3]).Value2 = "";
                        }
                    }
                }
                if (iRowStart + dt.Rows.Count - 1 - iSRowNO > 0)
                {
                    rg = sheet.get_Range(sheet.Cells[iSRowNO,1], sheet.Cells[iRowStart + dt.Rows.Count - 1,1]);
                    rg.MergeCells = true;
                    rg.WrapText = true;
                    rg = sheet.get_Range(sheet.Cells[iSRowNO,3], sheet.Cells[iRowStart + dt.Rows.Count - 1,3]);
                    rg.MergeCells = true;
                    rg.WrapText = true;
                }
                //�������ĺϼ� 2007-1-18 ���γ���ĺϼƣ�����Excel�Ĺ�ʽ����
                //Ҫ��д�������¼���������������Excel����ĺϼƹ�ʽ
                /*
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 4]).Value2 = "�ϼ�";
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 5]).Value2 = dbZCJF.ToString();
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 6]).Value2 = dbZCDF.ToString();
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 7]).Value2 = dbLRJF.ToString();
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 8]).Value2 = dbLRDF.ToString();
                 * */
            }
            catch
            {
                throw;
            }
        }
        #endregion

        #region liyuan 2006-12-28 д������ƽ���
        public void WriteTryCalcData(string strProID, string strRptNameID)
        {
            DbOperCls oper = new DbOperCls();
            string strSql = string.Empty;
            string strYear = string.Empty;
            string strRptType = "�걨";
            string strPeriod = "12";
            string strEntityID = "-1";
            int iTotalRow= 100;//ѭ�����ұ�����Ŀ������
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                oper.DbConnect();
                //������ĿID�õ����,��ֹ�գ������ֹ��<=6��ȡ���걨������ȡ�걨
                strSql = "select AccBeginYear,AccEndYear,DeadLine,EntityID from pro_project where id='" + strProID.ToString() + "'";
                ds = oper.GetSqlDataSet(strSql);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0][0].ToString() != "")
                    {
                        strYear = ds.Tables[0].Rows[0][0].ToString();
                    }
                    else if (ds.Tables[0].Rows[0][1].ToString() != "")
                    {
                        strYear = ds.Tables[0].Rows[0][1].ToString();
                    }
                    if (ds.Tables[0].Rows[0][2].ToString() != "")
                    {
                        if (DateTime.Parse(ds.Tables[0].Rows[0][2].ToString()).Month <= 6)
                        {
                            strRptType = "���걨";
                            strPeriod = "6";
                        }
                    }
                    if(ds.Tables[0].Rows[0][3].ToString()!="")
                        strEntityID = ds.Tables[0].Rows[0][3].ToString();
                }
                //����Excel�����ܱ�־Ϊ"XM",Ȼ��Ӹõ�Ԫ����һ����Ϊ�յ�Ԫ��ʼ��������
                Excel.Worksheet sheet = (Excel.Worksheet)m_XlApp.ActiveSheet;
                object oName = "XM";
                //??�˴�Names����Ӧ��ȡ��sheet,��֪Ϊ��ȥ����,��������applicationȡnames,���ܻ�������
                Excel.Name xmname = m_XlApp.Names.Item(oName, Optional, Optional);
                Excel.Range rg = null;
                int iBeginRow = 1;
                int iXMColumn = 1;
                int iColonIndex = -1;//ð�ŵ�����
                string strItemName = string.Empty;
                if(xmname != null)
                {
                    rg = xmname.RefersToRange;
                    iBeginRow = rg.Row+1;
                    iXMColumn = rg.Column;
                    for (int i = iBeginRow; i < iBeginRow + iTotalRow; i++)
                    {
                        iColonIndex = -1;
                        strItemName = ((Excel.Range)sheet.Cells[i, 1]).Text.ToString().Trim();
                        //ȥ����ǰ����ַ�
                        if (strItemName.IndexOf("��") != -1)
                        {
                            strItemName = strItemName.Substring(strItemName.IndexOf("��") + 1); 
                        }
                        //ȥ����ǰ����ַ�
                        if (strItemName.IndexOf(":") != -1)
                        {
                            iColonIndex = strItemName.IndexOf(":");
                        }
                        else
                        {
                            iColonIndex = strItemName.IndexOf("��");
                        }
                        if (iColonIndex != -1)
                        {
                            strItemName = strItemName.Substring(iColonIndex + 1);
                        }
                        if (strItemName.Trim() != "")
                        {
                            //���������ܱ�־ȡֵ
                            foreach (Excel.Name temp in m_XlApp.Names)
                            {
                                if (temp.Name.ToUpper() == "NC" || temp.Name.ToUpper() == "NM" || temp.Name.ToUpper() == "SN" || temp.Name.ToUpper() == "BN")
                                        ((Excel.Range)sheet.Cells[i, temp.RefersToRange.Column]).Value2 = Pub_Function.GetSSRtpValue(oper,strEntityID, strYear, strPeriod, strRptNameID, strItemName, temp.Name, "δ��");
                                    if (temp.Name.ToUpper() == "TZJF" || temp.Name.ToUpper() == "TZDF" || temp.Name.ToUpper() == "CFLJF" || temp.Name.ToUpper() == "CFLDF")
                                        ((Excel.Range)sheet.Cells[i, temp.RefersToRange.Column]).Value2 = Pub_Function.GetSSAdjustValue(oper,strProID, strYear, strRptType,strRptNameID, strItemName, temp.Name);
                            }
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                oper.DbClose();
            }
        }
        #endregion

        #region IDisposable ��Ա

        public void Dispose()
        {
            try
            {
                _startQuit = true;
                m_XlApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("����COM����������: " + ex.Message);
            }
            finally
            {
                OfficeWrap.NAR(m_XlApp);
                GC.Collect();
            }
        }

        #endregion
    }
}