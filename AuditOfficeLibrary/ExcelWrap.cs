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
        private bool _isAsyncSave = true;   //默认执行异步保存
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
                            //标记不存在,不做处理
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = "读取文档内的标志出现错误: " + ex.Message;
                    Debug.WriteLine(errMsg);
                    throw new Exception(errMsg, ex);
                }

                return marks;
            }
        }

        #region Obsolete! 原附注标志
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
                            //访问无效的Name的RefersToRange属性时会抛出异常,
                            //此处决定不加入"附注标志"字典
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
                    //如果标志的首字母不是大写的英文字母,则认为该标志属于"其它"类型
                    //todo:不要使用"汉字开头"作为"其它标志"的特征.
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
                                //访问无效的Name的RefersToRange属性时会抛出异常,
                                //此处决定不加入"其它标志"字典
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
                    //如果标志的首字母不是大写的英文字母,则认为该标志属于"其它"类型
                    if (nm.Name[0] > 90 || nm.Name[0] < 65)
                    {
                        otherMarkDictionary.Add(nm.Name, nm.RefersToRange.Text.ToString());
                    }
                }

                return otherMarkDictionary;
            }
        }

        /// <summary>
        /// 自定义,用来激发DocumentAfterSave事件
        /// </summary>
        protected virtual void OnDocumentAfterSave(AfterSaveEventArgs e)
        {
            if (this.DocumentAfterSaveEvent != null)
            {
                DocumentAfterSaveEvent(this, e);
            }
        }

        /// <summary>
        /// 自定义,用来激发DocumentBeforeSaveEvent事件
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnDocumentBeforeSave(BeforeSaveEventArgs e)
        {
            if (this.DocumentBeforeSaveEvent != null)
            {
                DocumentBeforeSaveEvent(this, e);
            }
        }

        #region 构造

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

            #region 为Excel Applicaton对象注册事件处理函数
            m_XlApp.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(_app_WorkbookBeforeSave);

            //关闭文档前
            m_XlApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(_app_WorkbookBeforeClose);
            #endregion

            //添加工具条
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
        /// 添加CommandBar
        /// </summary>
        protected override void AddCommandBar()
        {
            try
            {
                //为add-in建立一个命令条
                _commandBarSBMMark = m_XlApp.CommandBars.Add("智能标志",
                        MsoBarPosition.msoBarFloating,
                        Optional,
                        true);
                _commandBarSBMMark.Width = 100;

                base.AddCommandBarButtons();
            }
            catch (Exception ex)
            {
                string errMsg = "添加工具栏出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #endregion

        #region Event Handle Method

        private int _WbCount = -1;
        /// <summary>
        /// 在此处理Excel关闭问题
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

            //保证在Excel打开多个Workbook时,不会因为关闭其中的一个而关闭其它.
            if (_WbCount > 1)
            {
                isCloseApplication = false;
            }

            if (Path.GetExtension(Wb.Name) == ".txt")
            {
                Wb.Saved = true;    //让为全文检索准备的临时文本文件被忽略执行再保存
                return;
            }

            //无论是否修改过,都弹出"是否保存"的对话框.
            //ShowMessageBoxCallback showMessageBoxCallback = new ShowMessageBoxCallback(MessageBox.Show);
            //DialogResult result = (DialogResult)_excelFormThreadOwner.Invoke(showMessageBoxCallback, new object[] { String.Format("是否保存对{0}的更改?", Wb.Name), "Micorsoft Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation });
            if (!_startQuit)
            {
                DialogResult result;
                string text = String.Format("是否保存对{0}的更改?", Wb.Name);
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
                //Common.ShowMessageBox(String.Format("是否保存对{0}的更改?", Wb.Name),
                //    "确认",
                //    MessageBoxButtons.YesNo,
                //    _excelFormThreadOwner,
                //    out result);
                //Interaction.AppActivate("Microsoft Excel");

                switch (result)
                {
                    case DialogResult.Yes:
                        _isAsyncSave = false;
                        Wb.Save();
                        //wxg 2007年1月29日 10:58
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
            //todo:需要判断Wb是否是托管文档
            if (_isAsyncSave && !SaveAsUI)
            {
                OnDocumentBeforeSave(new BeforeSaveEventArgs(_docWrap));

                try
                {
                    //异步执行保存
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
                    string errMsg = "BackroundWorker出现错误:" + ex.Message;
                    //MessageBox.Show(errMsg);
                    Debug.WriteLine(errMsg);
                    throw new Exception(errMsg, ex);
                }
            }
        }

        /// <summary>
        /// 在此激发DocumentAfterSave事件
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
            //潜在的bug:如果用户在程序打开的Excel中新建、打开文档（不止xls）,
            //则保存此文档时，会将程序新建/打开文档时传入的DocWrap对象的Doc属性修改为此文档，
            //从而，实际被存入数据库的字节被偷换！    
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

        #region 文档操作
        /*
        public void OpenDoc(string filePathWithoutExtension, int id)
        {
            OpenDoc(filePathWithoutExtension);
            _docWrap.ID = id;
        }*/

        //TODO:NOTE:以后修改为私有或删除,目前外部仅为底稿管理中引入底稿使用(暗操作)
        /// <summary>
        /// 打开文档
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
                    //最大化
                    m_XlApp.WindowState = Excel.XlWindowState.xlMaximized;
                }

                Debug.WriteLine("文档操作:活动的Excel Workbook是: " + m_XlApp.ActiveWorkbook.Name);
            }
            catch (Exception ex)
            {
                string errMsg = "打开Workbook出现错误: " + ex.Message;
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
                    //最大化
                    m_XlApp.WindowState = Excel.XlWindowState.xlMaximized;
                }

                Debug.WriteLine("文档操作:活动的Excel Workbook是: " + m_XlApp.ActiveWorkbook.Name);
            }
            catch (Exception ex)
            {
                string errMsg = "打开Workbook出现错误: " + ex.Message;
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
                string errMsg = "新建Workbook出现错误: " + ex.Message;
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
                string errMsg = "新建Workbook出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        /// <summary>
        /// 保存文档(仅供新建时使用)
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
                string errMsg = "新建的Workbook保存到Temp文件夹失败: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #endregion

        #region CommandBarButton Event Handler
        protected override void btnInsertIndex_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //TODO:如果不判断是否托管文档,则在非托管文档中插入时,会查到相应的托管文档上?
            if (_docWrap.Doc == null) return;

            try
            {
                object address = null;
                object textToDisplay = null;
                BeforeInsertIndexEventArgs e = new BeforeInsertIndexEventArgs(address, textToDisplay);
                //激发插入索引事件
                OnInsertIndex(e);

                if (e.Address != null && e.TextToDisplay != null)
                {
                    address = e.Address.ToString();
                    textToDisplay = e.TextToDisplay;
                    //这里与Word的操作不同
                    ((Excel._Worksheet)(((Excel.Workbook)_docWrap.Doc).ActiveSheet)).Hyperlinks.Add(m_XlApp.ActiveCell, address.ToString(), Optional, Optional, textToDisplay);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "插入交叉索引出现错误: " + ex.Message;
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
                string errMsg = "保存标记出现错误: " + ex.Message;
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
                string errMsg = "保存附注出现错误: " + ex.Message;
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

        //NOTE!仅在保存后处理时调用.
        public override string GetDocText(string filePath)
        {
            List<string> txtFileNameCollection = new List<string>();
            StringBuilder sb = new StringBuilder();

            try
            {
                //using (ExcelWrap excelWrap = new ExcelWrap(false, null))
                //{
                //    //NOTE!!确保DoWorkAfterSaved方法不被调用!
                //    excelWrap.IsAsyncSave = false;

                //    string fileFullName = OfficeWrap.GetFullName(filePath);
                //    excelWrap.OpenDoc(fileFullName);
                //    Excel.Workbook xlBook = (Excel.Workbook)excelWrap.DocWrap.Doc;

                //    //以CSV格式读取Workbook中的文本
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
                    //确保DoWorkAfterSaved方法不被调用.
                    excelWrap.IsAsyncSave = false;
                    excelWrap.OpenDoc(tempFilePath);
                    Excel.Workbook xlBook = (Excel.Workbook)excelWrap.DocWrap.Doc;
                    string txtFileName;
                    foreach (Excel.Worksheet xlSheet in xlBook.Worksheets)
                    {
                        txtFileName = Common.GetTempDirectoryPath() + Common.NewId() + ".txt";
                        //以CSV格式读取Workbook中的文本
                        xlSheet.SaveAs(txtFileName, Excel.XlFileFormat.xlCSV,
                            Optional, Optional, Optional, Optional, Optional, Optional, Optional);

                        txtFileNameCollection.Add(txtFileName);
                    }
                }

                foreach (string fileName in txtFileNameCollection)
                {
                    using (FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        TextReader reader = new StreamReader(stream, Encoding.Default);//Encoding采用了试验的结果
                        sb.Append(reader.ReadToEnd());
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "获取全文本出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return sb.ToString();
        }

        //NOTE!仅在保存后处理时调用.
        public override List<string> GetCrossRefList()
        {
            List<string> crossRefIdCollection = new List<string>();
            Excel.Workbook wb = (Excel.Workbook)_docWrap.Doc;

            if (!(Path.GetExtension(wb.Name) == ".txt"))//如果是其它类型的文件,Excel的相应是什么样的?
            {
                try
                {
                    foreach (Excel.Hyperlink hl in ((Excel._Worksheet)wb.ActiveSheet).Hyperlinks)
                    {
                        string strRefedDocId = Path.GetFileNameWithoutExtension(hl.Address);
                        int refedDocId;
                        if (Int32.TryParse(strRefedDocId, out refedDocId))
                        {
                            //TODO:对哪些链接的地址是引用到其它底稿作判断,以保持此方法的通用性
                            //不保存重复的交叉索引,以防止在文档打开过程中下载引用的文件时,
                            //重复下载同一文件,并导致IO异常;
                            //不保存对文档自身的引用,理由同上;
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
                    string errMsg = "获取交叉索引出现错误: " + ex.Message;
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
                        if (value2.EndsWith(":") || value2.EndsWith("："))
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
                //由于_excelSigleton.ActiveCell.Name 不为null,
                //无法判断Cell的Range是否有Name,所以无论有无,强行删除.
            } 
        }

        /// <summary>
        /// 在Excel的ActiveCell上插入NamedRange
        /// 如果是报表或余额表标志,则先清除原来的NamedRange后再插入,
        /// </summary>
        /// <param name="markName">如果是String.Empty,则不做处理</param>
        public override void UpdateMark(string markName)
        {
            try
            {
                if(markName != "")
                {
                    Excel.Name nme = m_XlApp.ActiveWorkbook.Names.Add(markName, m_XlApp.ActiveCell, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional, OfficeWrap.Optional);

                    nme.Visible = true;//作用?
                }
            }
            catch (Exception ex)
            {
                string errMsg = "插入标志出现错误: " + ex.Message;
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
            //删除标记未替换时表示标记的"<mark>"文本
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

                        int colonIndex = value.Contains(":") ? value.IndexOf(':') : value.IndexOf('：');
                        string iniValue = value.Substring(0, colonIndex + 1);

                        //保证插入时间字符串时,Excel不会将单元格的格式自动设置为时间类型,而是保持文本格式
                        //rng.NumberFormatLocal = "＠";
                        rng.Value2 = iniValue + mark.Value;
                    //    break;
                    //}
                }
            }
            catch (Exception ex)
            {
                string errMsg = "替换标志时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        #region 处理调整分录标记
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
                            xlStartRange1.Value2 = "重分类分录：";
                            xlStartRange2 = xlStartRange1.get_Offset(adTableList[0].Rows.Count + 1, 0);
                            Foo(adTableList[0], xlStartRange1);
                        }
                        else
                        {
                            xlStartRange2 = xlStartRange1;
                        }

                        if (adTableList[1].Rows.Count > 0)
                        {
                            xlStartRange2.Value2 = "调整分录：";
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
                if (dr[0].ToString().StartsWith("借"))  //borrow
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

            DataColumn colReportItem    = new DataColumn("ReportItem"); //报表项
            DataColumn colXM            = new DataColumn(Ano.AnoXM); //项
            DataColumn colNC            = new DataColumn(Ano.AnoNC, typeof(double));   //年初数
            DataColumn colNM            = new DataColumn(Ano.AnoNM, typeof(double));   //年末数
            DataColumn colJL            = new DataColumn(Ano.AnoJL, typeof(double));   //增加
            DataColumn colDL            = new DataColumn(Ano.AnoDL, typeof(double));   //减少
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
                                //访问无效的Name的RefersToRange属性时会抛出异常,
                                //此处决定不加入"附注标志"字典
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

            //不存在指定的附注标志,立即返回
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
        /// 判断打开的Excel实例中是否存在指定的名称
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
                string errMsg = "检查NamedRange是否存在时出现错误:" + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return existName;
        }

        //特殊粘贴
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
                string errMsg = "粘贴时出现错误: " + ex.Message;
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
                //低效
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
                //从同一个Excel中复制粘贴
                //nme.RefersToRange.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Optional, Optional);

                //从不同的Excel总复制粘贴
                //ActiveSheet.PasteSpecial Format:="文本", Link:=False, DisplayAsIcon:=False
                nme.RefersToRange.Worksheet.Activate();
                nme.RefersToRange.Select();
                ((Excel.Worksheet)nme.Application.ActiveSheet).PasteSpecial("文本", False, False, Optional, Optional, Optional);
            }
            catch (Exception)
            {

                throw;
            }

        }

        #region liyuan 2006-12-27 写入差异数据
        //查找第一列不是单元格的行,然后从下二行开始写入数据,考虑合并单元格的问题,所以
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
                            //1:资产负债,2:利润
                            if (dt.Rows[i][arrColumn[j]].ToString().Trim() == "1")
                            {
                                //资产负债表
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
                                //利润表
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
                //填充数据完毕后,对分录号,调整原因列设置合并单元格
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
                //加入最后的合计 2007-1-18 屏蔽程序的合计，采用Excel的公式定义
                //要求写入调整分录的数量，否则会冲掉Excel定义的合计公式
                /*
                ((Excel.Range)sheet.Cells[iRowStart + dt.Rows.Count, 4]).Value2 = "合计";
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

        #region liyuan 2006-12-28 写入试算平衡表
        public void WriteTryCalcData(string strProID, string strRptNameID)
        {
            DbOperCls oper = new DbOperCls();
            string strSql = string.Empty;
            string strYear = string.Empty;
            string strRptType = "年报";
            string strPeriod = "12";
            string strEntityID = "-1";
            int iTotalRow= 100;//循环查找报表项目的行数
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                oper.DbConnect();
                //根据项目ID得到年度,截止日，如果截止日<=6则取半年报，否则取年报
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
                            strRptType = "半年报";
                            strPeriod = "6";
                        }
                    }
                    if(ds.Tables[0].Rows[0][3].ToString()!="")
                        strEntityID = ds.Tables[0].Rows[0][3].ToString();
                }
                //查找Excel中智能标志为"XM",然后从该单元格下一个不为空单元格开始计算数据
                Excel.Worksheet sheet = (Excel.Worksheet)m_XlApp.ActiveSheet;
                object oName = "XM";
                //??此处Names集合应该取自sheet,不知为何去不到,所以先用application取names,可能会有问题
                Excel.Name xmname = m_XlApp.Names.Item(oName, Optional, Optional);
                Excel.Range rg = null;
                int iBeginRow = 1;
                int iXMColumn = 1;
                int iColonIndex = -1;//冒号的索引
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
                        //去掉、前面的字符
                        if (strItemName.IndexOf("、") != -1)
                        {
                            strItemName = strItemName.Substring(strItemName.IndexOf("、") + 1); 
                        }
                        //去掉：前面的字符
                        if (strItemName.IndexOf(":") != -1)
                        {
                            iColonIndex = strItemName.IndexOf(":");
                        }
                        else
                        {
                            iColonIndex = strItemName.IndexOf("：");
                        }
                        if (iColonIndex != -1)
                        {
                            strItemName = strItemName.Substring(iColonIndex + 1);
                        }
                        if (strItemName.Trim() != "")
                        {
                            //许还数据智能标志取值
                            foreach (Excel.Name temp in m_XlApp.Names)
                            {
                                if (temp.Name.ToUpper() == "NC" || temp.Name.ToUpper() == "NM" || temp.Name.ToUpper() == "SN" || temp.Name.ToUpper() == "BN")
                                        ((Excel.Range)sheet.Cells[i, temp.RefersToRange.Column]).Value2 = Pub_Function.GetSSRtpValue(oper,strEntityID, strYear, strPeriod, strRptNameID, strItemName, temp.Name, "未审");
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

        #region IDisposable 成员

        public void Dispose()
        {
            try
            {
                _startQuit = true;
                m_XlApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("结束COM服务器错误: " + ex.Message);
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