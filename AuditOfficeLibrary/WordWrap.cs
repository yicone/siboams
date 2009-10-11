using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using Word;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Office;
using AuditPubLib;
using System.Data;
using Microsoft.VisualBasic;

namespace AuditOfficeLibrary
{
    public class WordWrap : OfficeWrap, IDisposable
    {
        public event DocumentBeforeSaveHandler DocumentBeforeSaveEvent;
        public event DocumentAfterSaveHandler DocumentAfterSaveEvent;


        private Word.Application m_WdApp = null;
        private Form _wordFormThreadOwner = null;
        private Word.Document _activeDoc = null;
        private BackgroundWorker _saveWorker = new BackgroundWorker();
        //用于读取文档的文本内容时，是否直接读取
        private bool _documentWillClose = false;        
        private object _readonly = false;
        private DocWrap _docWrap = null;


        public override DocWrap DocWrap
        {
            get { return _docWrap; }
        }

        public bool Visable
        {
            get { return m_WdApp.Visible; }
            set { m_WdApp.Visible = value; }
        }

        //NOTE:set此属性时并不直接应用于Word
        public bool Readonly
        {
            get { return (bool)_readonly; }
            set { _readonly = value; }
        }

        //返回所有书签,不区分书签是否由用户定义
        public Word.Bookmarks Bookmarks
        {
            get
            {
                return m_WdApp.ActiveDocument.Bookmarks;
            }
        }

        //返回所有书签的包装类Mark的集合
        public override List<Mark> Marks
        {
            get
            {
                List<Mark> marks = new List<Mark>();
                Mark mark;
                try
                {
                    foreach (Word.Bookmark bmk in m_WdApp.ActiveDocument.Bookmarks)
                    {
                        mark = new Mark(bmk.Name, bmk.Start, bmk.End, -1);
                        marks.Add(mark);
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
                foreach (Word.Bookmark bmk in _wordSingleton.ActiveDocument.Bookmarks)
                {
                    if (bmk.Name.ToUpper().StartsWith("ANNO"))
                    {
                        annoDictionary.Add(bmk.Name, bmk.Range.Text.Replace("\r\a", ""));
                    }
                }
                return annoDictionary;
            }
        } */
        #endregion

        //返回文档中所有"其它"源标志的字典
        public override Dictionary<string, string> RefedMarkDictionary
        {
            get
            {
                Dictionary<string, string> refedMarkDictionary = new Dictionary<string, string>();
                foreach (Word.Bookmark bmk in m_WdApp.ActiveDocument.Bookmarks)
                {
                    //如果标志的首字母不是大写的英文字母,则认为该标志属于"其它"类型
                    if (bmk.Name[0] > 90 || bmk.Name[0] < 65)
                    {
                        if (bmk.Name.EndsWith("_"))
                        {
                            refedMarkDictionary.Add(bmk.Name, bmk.Range.Text.Replace("\r\a", ""));
                        }
                    }
                }

                return refedMarkDictionary;
            }
        }

        //返回文档中所有"其它"值标志的字典
        public override Dictionary<string, string> OtherMarkDictionary
        {
            get
            {
                Dictionary<string, string> otherMarkDictionary = new Dictionary<string, string>();
                foreach (Word.Bookmark bmk in m_WdApp.ActiveDocument.Bookmarks)
                {
                    //如果标志的首字母不是大写的英文字母,则认为该标志属于"其它"类型
                    if (bmk.Name[0] > 90 || bmk.Name[0] < 65)
                    {
                        otherMarkDictionary.Add(bmk.Name, bmk.Range.Text.Replace("\r\a", ""));
                    }
                }

                return otherMarkDictionary;
            }
        }


        //构造函数
        public WordWrap(bool visable, Form wordFormThreadOwner)
        {
            _wordFormThreadOwner = wordFormThreadOwner;
            _saveWorker.WorkerSupportsCancellation = true;
            _saveWorker.DoWork += new DoWorkEventHandler(_saveWorker_DoWork);
            _saveWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_saveWorker_RunWorkerCompleted);

            m_WdApp = new Word.Application();
            m_WdApp.Visible = visable;

            //添加工具条
            AddCommandBar();

            #region 为Word Applicaton对象注册保存前,关变前,及应用程序退出事件的处理函数
            m_WdApp.DocumentBeforeSave += new ApplicationEvents2_DocumentBeforeSaveEventHandler(_word_DocumentBeforeSave);

            ((Word.ApplicationClass)m_WdApp).ApplicationEvents2_Event_Quit += new ApplicationEvents2_QuitEventHandler(WordWrap_ApplicationEvents2_Event_Quit);

            m_WdApp.DocumentBeforeClose += new ApplicationEvents2_DocumentBeforeCloseEventHandler(_word_DocumentBeforeClose);
            #endregion

            try
            {
                //NOTE:确保不修改到Normal.dot
                object fileName = "Normal.dot";
                //_word.Templates.get_Item(ref fileName).Saved = true;
                m_WdApp.Templates.Item(ref fileName).Saved = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Operate 'Noraml.dot' Exception: " + ex.Message);
            }
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


        #region 智能标志操作
        protected override void btnDeleteMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                DeleteMarkForm frmDeleteMark = new DeleteMarkForm(this);

                _wordFormThreadOwner.Invoke(new MethodInvoker(frmDeleteMark.Show));
                _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.SendToBack));
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

                _wordFormThreadOwner.Invoke(new MethodInvoker(frmInsertMark.Show));
                _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.SendToBack));
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
                m_WdApp.Selection.SelectCell();
                //Word.Cell cell = _word.Selection.Cells[1];
                //office2000:
                Word.Cell cell = m_WdApp.Selection.Cells.Item(1);

                if (cell == null) return;

                int rowIndex = cell.RowIndex;
                int colIndex = cell.ColumnIndex;
                string mark = "";
                if (colIndex - 1 >= 0)
                {
                    try
                    {
                        mark = cell.Previous.Range.Text.Replace("\r\a", "");
                    }
                    catch { }
                    //string markValue = cell.Range.Text.Replace("\r\a", "");
                }

                InsertOtherMarkForm frm = new InsertOtherMarkForm(this, mark);
                _wordFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.SendToBack));
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
                //_wordSingleton.Selection.SelectCell();
                //Word.Cell cell = _wordSingleton.Selection.Cells.Item(1);
                //if (cell == null) return;

                //int rowIndex = cell.RowIndex;
                //int colIndex = cell.ColumnIndex;
                //if (colIndex - 1 >= 0)
                //{
                //    string mark = cell.Previous.Range.Text.Replace("\r\a", "");
                //    string markValue = cell.Range.Text.Replace("\r\a", "");

                InsertAnoForm frm = new InsertAnoForm(this);
                _wordFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.SendToBack));
                //}
            }
            catch (Exception ex)
            {
                string errMsg = "保存附注出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }
        #endregion

        //插入"交叉索引"
        protected override void btnInsertIndex_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                //object textToDisplay = "Test Insert Index";
                //object address = "http://www.microsoft.com";
                //_activeDoc.Hyperlinks.Add(_word.Selection.Range, ref address, ref OfficeWrap.Optional, ref OfficeWrap.Optional, ref address, ref OfficeWrap.Optional);

                object textToDisplay = null;
                object address = null;
                BeforeInsertIndexEventArgs e = new BeforeInsertIndexEventArgs(address, textToDisplay);
                //激发插入索引事件
                OnInsertIndex(e);

                if (e.Address != null && e.TextToDisplay != null)
                {
                    address = e.Address.ToString();
                    textToDisplay = e.TextToDisplay;
                    _activeDoc.Hyperlinks.Add(m_WdApp.Selection.Range, ref address, ref Optional, ref Optional, ref textToDisplay, ref Optional);
                    m_WdApp.Activate();
                }
            }
            catch (Exception ex)
            {
                string errMsg = "插入交叉索引出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        //插入审计结论
        protected override void btnInsertResult_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                InsertResultForm frm = new InsertResultForm(this);
                _wordFormThreadOwner.Invoke(new MethodInvoker(frm.Show));
                _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.SendToBack));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        } 

        private void _word_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            //_activeDocFileName = Path.Combine(DocWrap.Path, DocWrap.Name);
            _documentWillClose = true;
            ////wxg 2007年1月29日 15:58
            //OnDocumentBeforeSave(new BeforeSaveEventArgs(_docWrap));
        }

#if DEBUG
        private int m_ThreadNum = 0;
#endif
        private void _word_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            if (Doc.Name == "Normal.dot")
                return;

            OnDocumentBeforeSave(new BeforeSaveEventArgs(_docWrap));

            try
            {
                if (_saveWorker.IsBusy)
                {
                    _saveWorker.CancelAsync();
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
                }

                _saveWorker.RunWorkerAsync(Doc);
            }
            catch (Exception ex)
            {
                string errMsg = "BackroundWorker出现错误:" + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        private void WordWrap_ApplicationEvents2_Event_Quit()
        {
            try
            {
                m_WdApp.CommandBars[_commandBarSBMMark.Name].Delete();
            }
            catch (Exception ex)
            {
                string errMsg = "退出Word时删除工具栏出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            OfficeWrap.NAR(_commandBarSBMMark);
            //关闭Word后,焦点回到主窗口
            _wordFormThreadOwner.Invoke(new MethodInvoker(_wordFormThreadOwner.Activate));
        }

        //异步操作:执行保存
        private void _saveWorker_DoWork(object sender, DoWorkEventArgs e)
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
            IntPtr p = Common.FindWindow(null, "另存为"); 
            if (p != IntPtr.Zero)
            {
                e.Cancel = true;
#if DEBUG
                MessageBox.Show("抓到另存为了");
#endif
                return;
            }


            try
            {
                _activeDoc = e.Argument as Word.Document;

                int waitTime = 0;
                while (!_activeDoc.Saved)
                {
                    Thread.Sleep(100);
                    waitTime += 100;
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
                Thread.Sleep(100);  //不确切
                OnDocumentAfterSave(new AfterSaveEventArgs(_docWrap));
                return;
            }
            catch (COMException)
            {
                Thread.Sleep(100);  //不确切
                OnDocumentAfterSave(new AfterSaveEventArgs(_docWrap));
                return;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                MessageBox.Show("调用OnDocumentAfterSave方法前出错");
            }
        }

        #region 文档操作:新建、打开
        /*
        public void OpenDoc(string filePathWithoutExtension, int id)
        {
            OpenDoc(filePathWithoutExtension);
            _docWrap.ID = id;
        }*/

        public void OpenDoc(string filePathWithoutExtension)
        {
            object fileName = filePathWithoutExtension;
            object visable = m_WdApp.Visible;

            try
            {
//#if OFFICE2000
//            _activeDoc = _word.Documents.Open2000(ref fileName,
//                                                ref Optional,
//                                                ref _readonly,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref Optional,
//                                                ref visable);
//#else
//                _activeDoc = _word.Documents.Open(ref fileName,
//                    ref Optional,
//                    ref _readonly,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref visable,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional);
//#endif
                //2000:12
                _activeDoc = m_WdApp.Documents.Open(ref fileName,
                    ref Optional,
                    ref _readonly,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref visable);

                if (m_WdApp.Visible)
                {
                    m_WdApp.Activate();
                }

                _docWrap = new DocWrap();
                _docWrap.Doc = _activeDoc;
                _docWrap.EditState = 2;
                _docWrap.Path = filePathWithoutExtension;
            }
            catch (Exception ex)
            {
                string errMsg = "打开Document出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void OpenDoc(DocWrap docWrap)
        {
            _docWrap = docWrap;
            object fileName = _docWrap.Path;
            object visable = m_WdApp.Visible;

            try
            {
                //2000:12
                _activeDoc = m_WdApp.Documents.Open(ref fileName,
                    ref Optional,
                    ref _readonly,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref visable);

                if (m_WdApp.Visible)
                {
                    m_WdApp.Activate();
                }

                _docWrap.Doc = _activeDoc;
                _docWrap.EditState = 2;
            }
            catch (Exception ex)
            {
                string errMsg = "打开Document出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                if(ex is COMException && ((COMException)ex).ErrorCode != -2146824090)
                    throw new Exception(errMsg, ex);
            }
        }

        /*
        public void NewDoc(int docID, string filePathWithoutExtension, bool immediatelySave)
        {
            object visable = _wordSingleton.Visible;
            
            try
            {
#if OFFICE2000
                _activeDoc = _wordSingleton.Documents.AddOld(ref OfficeWrap.Optional, ref OfficeWrap.Optional);
#else
                _activeDoc = _word.Documents.Add(ref Optional,
                                       ref Optional,
                                       ref Optional,
                                       ref visable);
#endif

                _docWrap.ID = docID;
                _docWrap.Doc = _activeDoc;
                _docWrap.EditState = 0;//New
                _docWrap.Path = filePathWithoutExtension;
                Debug.WriteLine(String.Format("文档操作:新建文件{0}路径及状态初始完成", filePathWithoutExtension));

                Thread.Sleep(300);
                if (immediatelySave)
                {
                    SaveDocToLocal(filePathWithoutExtension);
                }

                if (_wordSingleton.Visible)
                {
                    _wordSingleton.Activate();
                }
            }
            catch (Exception ex)
            {
                string errMsg = "新建Document出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }*/

        public void NewDoc(DocWrap docWrap, bool immediatelySave)
        {
            _docWrap = docWrap;
            try
            {
                object visible = m_WdApp.Visible;
                _activeDoc  = m_WdApp.Documents.Add(ref Optional, ref Optional, ref Optional, ref visible);
                _docWrap.Doc = _activeDoc;
                _docWrap.EditState = 0;//New

                Debug.WriteLine(String.Format("文档操作:新建文件{0}路径及状态初始完成", docWrap.Path));

                Thread.Sleep(300);
                if (immediatelySave)
                {
                    SaveDocToLocal(docWrap.Path);
                }

                if (m_WdApp.Visible)
                {
                    m_WdApp.Activate();
                }
            }
            catch (Exception ex)
            {
                string errMsg = "新建Document出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        private void SaveDocToLocal(string filePathWithoutExtension)
        {
            object fileName = filePathWithoutExtension;

            try
            {
//#if OFFICE2000
//            _activeDoc.SaveAs2000(ref fileName,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional,
//                ref Optional);
//#else
//                _activeDoc.SaveAs(ref fileName,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional,
//                    ref Optional);
//#endif
                //2000:12s
                _activeDoc.SaveAs(ref fileName,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional,
                    ref Optional);

            }
            catch (Exception ex)
            {
                string errMsg = "新建的Document保存到Temp文件夹出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }
        #endregion

        //取得活动文档内容的字节数组
        public override Byte[] GetDocBytes(string fileFullName)
        {
            string filePath = fileFullName + ".doc";
            return base.GetDocBytes(filePath);
        }

        //取得指定文件的Text. NOTE:可能需要创建新的Word对象!
        public override string GetDocText(string filePath)
        {
            string text = "";

            try
            {
                //文件没有关闭,直接读取Text;
                //若关闭,需要新建一个Word来完成操作
                //if (!_documentWillClose)
                //{
                    text = m_WdApp.ActiveDocument.Content.Text;
                //}
                //else
                //{
                    //using (WordWrap wordWrap = new WordWrap(false, null))
                    //{
                    //    string fileFullName = OfficeWrap.GetFullName(filePath);
                    //    wordWrap.OpenDoc(fileFullName);
                    //    //todo:不要出现_activeDoc.
                    //    text = ((Word.Document)wordWrap.DocWrap.Doc).Content.Text;
                    //}
                //}
            }
            catch (Exception ex)
            {
                string errMsg = "获取全文本出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return text;
        }

        //取得Word文档中的超链接集合
        public override List<string> GetCrossRefList()
        {
            List<string> crossRefIdCollection = new List<string>();
            try
            {
                foreach (Word.Hyperlink hl in _activeDoc.Hyperlinks)
                {
                    int temp;
                    string strRefId = Path.GetFileNameWithoutExtension(hl.Address);
                    if (Int32.TryParse(strRefId, out temp))
                    {
                        //TODO:对哪些链接的地址是引用到其它底稿作判断,以保持此方法的通用性
                        //NOTE:不保存重复的交叉索引;
                        //NOTE:不保存对文档自身的引用
                        //原因:防止在文档打开过程中下载引用的文件时,
                        //重复下载同一文件,并导致IO异常.
                        if (!crossRefIdCollection.Contains(strRefId) &&
                            strRefId != Path.GetFileNameWithoutExtension(_activeDoc.Name))
                        {
                            crossRefIdCollection.Add(strRefId);
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

            return crossRefIdCollection;
        }

        public override void AppendText(string text)
        {
            try
            {
                m_WdApp.Selection.TypeText(text);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// 在Word文档的选定位置插入书签
        /// 由于不涉及报表和余额表标志,无需考虑是否先清除书签后再插入
        /// </summary>
        /// <param name="mark"></param>
        public override void UpdateMark(string mark)
        {
            try
            {
                //if (!_wordSingleton.ActiveDocument.Bookmarks.Exists(mark))
                //{
                    //_wordSingleton.Selection.Text = String.Format(@"<{0}>", mark);
                    object rnge = m_WdApp.Selection.Range;
                    Word.Bookmark bmk = m_WdApp.ActiveDocument.Bookmarks.Add(mark, ref rnge);
                //}
                //else
                //{
                //    throw new Exception(string.Format(@"标志{0}已存在.", mark));
                //}
            }
            catch (Exception ex)
            {
                string errMsg = "插入标志出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public override void DeleteMark(string mark)
        {
            object oMark = mark;
            //删除标记未替换时表示标记的"<mark>"文本
            string bmkText = m_WdApp.ActiveDocument.Bookmarks.Item(ref oMark).Range.Text;

            if (bmkText != null && bmkText.StartsWith("<") && bmkText.EndsWith(">"))
            {
                //注意:清空文本的同时,书签已被删除
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

        //添加工具栏到Word中
        protected override void AddCommandBar()
        {
            try
            {
                //为add-in建立一个命令条
                _commandBarSBMMark = m_WdApp.CommandBars.Add("智能标志",
                        MsoBarPosition.msoBarFloating,
                        Optional,
                        true);
                _commandBarSBMMark.Width = 100;

                base.AddCommandBar();
            }
            catch (Exception ex)
            {
                string errMsg = "添加工具栏出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        //自定义,用于激发DocumentAfterSave事件
        protected virtual void OnDocumentAfterSave(AfterSaveEventArgs e)
        {
            if (this.DocumentAfterSaveEvent != null)
            {
                DocumentAfterSaveEvent(this, e);
            }
        }

        //自定义,用于激发DocumentBeforeSaveEvent事件
        protected virtual void OnDocumentBeforeSave(BeforeSaveEventArgs e)
        {
            if (this.DocumentBeforeSaveEvent != null)
            {
                DocumentBeforeSaveEvent(this, e);
            }
        }

        public void ReplaceAllMarks(List<Mark> marks)
        {
            try
            {
                //foreach (Mark mark in marks)
                //{
                //    object start = mark.RowNum;
                //    object end = mark.ColNum;
                //    Word.Range range = _wordSingleton.ActiveDocument.Range(ref start, ref end);
                //    string value = range.Text == null ? "" : range.Text;
                //    int colonIndex = value.Contains(":") ? value.IndexOf(':') : value.IndexOf('：');
                //    string iniValue = value.Substring(0, colonIndex + 1);
                //    //_wordSingleton.ActiveDocument.Range(ref start, ref end).Text = iniValue + mark.Value;
                //    //object oBookmark = 1;
                //    //if (range.Bookmarks.Count == 0) continue;
                //    //Word.Range rge = range.Bookmarks.Item(ref oBookmark).Range;
                //    range.InsertBefore(iniValue + mark.Value);
                //    ////重新插入书签,以解决替换标志时书签被删除的问题
                //    //end = Convert.ToInt32(start) + mark.Value.Length;
                //    //object bmkRng = _word.ActiveDocument.Range(ref start, ref end);
                //    //_word.ActiveDocument.Bookmarks.Add(mark.Formula, ref bmkRng);
                //}

                List<Word.Bookmark> bookmarkCollection = new List<Bookmark>();
                foreach (Word.Bookmark bookmark in m_WdApp.ActiveDocument.Bookmarks)
                {
                    bookmarkCollection.Add(bookmark);
                }

                foreach (Word.Bookmark bookmark in bookmarkCollection)
                {
                    foreach(Mark mark in marks)
                    {
                        if(mark.Formula == bookmark.Name)
                        {
                            Word.Range range = bookmark.Range;
                            if (mark.Value.Contains("|"))
                            {
                                range.Select();
                                string[] paragraphs = mark.Value.Split('|');
                                for (int i = 0; i < paragraphs.Length; i++ )
                                {
                                    string str = paragraphs[i];
                                    if (i != 0)
                                    {
                                        m_WdApp.Selection.TypeParagraph();
                                    }
                                    m_WdApp.Selection.TypeText(str);
                                }

                                range.SetRange(range.Start, m_WdApp.Selection.End);
                            }
                            else
                            {
                                range.Text = mark.Value;
                            }
                            range.Select();
                            object oRange = m_WdApp.Selection.Range;
                            m_WdApp.ActiveDocument.Bookmarks.Add(mark.Formula, ref oRange);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "替换标志时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                //throw new Exception(errMsg, ex);
            }
        }

        public void Paste()
        {
            try
            {
                //插到文档的有内容的最后一行的后面
                //方法一:
                //object end = _word.ActiveDocument.Content.End - 1 ;
                //_word.ActiveDocument.Range(ref end, ref OfficeWrap.Optional).Paste();

                //方法二:
                object units = Word.WdUnits.wdStory;
                m_WdApp.Selection.EndKey(ref units, ref Optional);
                m_WdApp.Selection.Paste();
            }
            catch (Exception ex)
            {
                string errMsg = "粘贴时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void InsertPageBreak()
        {
            try
            {
                if (m_WdApp.Selection != null)
                {
                    object oEnd = m_WdApp.Selection.End;
                    object oBreakType = WdBreakType.wdPageBreak;
                    Word.Range wdRng = m_WdApp.ActiveDocument.Range(ref oEnd, ref oEnd);
                    wdRng.InsertBreak(ref oBreakType);
                }
            }
            catch
            {
                throw;
            }
        }

        public void CopyAll()
        {
            try
            {
                m_WdApp.Selection.WholeStory();
                m_WdApp.Selection.Copy();
            }
            catch (Exception ex)
            {
                string errMsg = "复制时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void PasteBaseInfoWorksheet()
        {
            try
            {
                bool existsBaseInfoBookmark = m_WdApp.ActiveDocument.Bookmarks.Exists(ReplaceMark.BaseInfo);
                if (!existsBaseInfoBookmark) return;

                object oBaseInfo = ReplaceMark.BaseInfo;
                Word.Bookmark bmkBaseInfo = m_WdApp.ActiveDocument.Bookmarks.Item(ref oBaseInfo);
                Range range = bmkBaseInfo.Range;
                range.Paste();
                //range.InsertBefore("附件1       ");
                range.Select();
                foreach (Word.Bookmark bmk in m_WdApp.Selection.Bookmarks)
                {
                    bmk.Delete();
                }
                object oRange = m_WdApp.Selection.Range;
                m_WdApp.ActiveDocument.Bookmarks.Add(ReplaceMark.BaseInfo, ref oRange);
                ////换页
                //object wdPageBreak = WdBreakType.wdPageBreak;
                //range.InsertBreak(ref wdPageBreak);
            }
            catch (Exception ex)
            {
                string errMsg = "粘贴基本情况时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        public void ReplaceAnoMarks(int projectId)
        {
            Dictionary<string, int> anoColumnDictionary = null;

            try
            {
                Word.Bookmark bmkXM = null;
                Word.Cell cellXM = null;
                Word.Table table = null;

                foreach (Word.Bookmark bmk in m_WdApp.ActiveDocument.Bookmarks)
                {
                    anoColumnDictionary = new Dictionary<string, int>();
                   
                    cellXM = GetAnoCell(Ano.AnoXM, bmk);
                    if (cellXM != null)
                    {
                        bmkXM = bmk;
                        table = cellXM.Range.Tables.Item(1);
                        anoColumnDictionary.Add(Ano.AnoXM, cellXM.ColumnIndex);

                        #region 取附注所在的列号
                        foreach (Word.Bookmark bmk2 in table.Range.Bookmarks)
                        {
                            AddAnoColNumPair(bmk2, anoColumnDictionary);
                        } 
                        #endregion

                        //清除上次替换出的行
                        int i = table.Rows.Count;
                        while (i != 1)
                        {
                            table.Rows.Item(i).Delete();
                            i--;
                        }

                        string[] array = bmkXM.Name.Split('_');
                        string reportItem = array[2];
                        List<string> xmCollection = DAL.GetXMCollection(projectId, reportItem);

                        int x = 2;
                        foreach (string xmName in xmCollection)
                        {
                            Word.Row newRow = table.Rows.Add(ref Optional);

                            foreach (KeyValuePair<string, int> kvp in anoColumnDictionary)
                            {
                                if (kvp.Key == Ano.AnoXM)
                                {
                                    table.Cell(x, kvp.Value).Range.Text = xmName;
                                }
                                else
                                {
                                    Word.Cell cell = table.Cell(x, kvp.Value);
                                    string[] array1 = kvp.Key.Split('_');
                                    string anoCol = array1[0] + array1[1];
                                    cell.Range.Text = DAL.GetAnoValue(projectId, reportItem, xmName, anoCol);
                                    if (!array1[1].ToUpper().StartsWith("TXT"))
                                    ToRight(cell);
                                }
                            }

                            x++;
                        }

                        if (x > 2)
                        {
                            table.Rows.Add(ref Optional);

                            foreach (KeyValuePair<string, int> kvp in anoColumnDictionary)
                            {
                                if (kvp.Key == Ano.AnoXM)
                                {
                                    table.Cell(x + 1, kvp.Value).Range.Text = "合计";
                                }
                                else
                                {
                                    Word.Cell cell = table.Cell(x + 1, kvp.Value);
                                    string[] array1 = kvp.Key.Split('_');
                                    string anoCol = array1[0] + array1[1];
                                    cell.Range.Text = DAL.GetAnoSumValue(projectId, reportItem, anoCol);
                                    if(!array1[1].ToUpper().StartsWith("TXT"))
                                        ToRight(cell);
                                }
                            }
                        }//end if
                    }//end if
                }//end if
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void ReplaceBookmarkText(string bmkName, string bmkText)
        {
            try
            {
                //object oIndex = 1;
                //Word.Document wdDoc = m_WdApp.Documents.Item(ref oIndex);
                Word.Document wdDoc = m_WdApp.ActiveDocument;
                foreach (Word.Bookmark wdBmk in wdDoc.Bookmarks)
                {
                    if (wdBmk.Name == bmkName)
                    {
                        Word.Range wdRng = wdBmk.Range;
                        wdRng.Text = bmkText;
                        wdRng.Select();
                        object oRange = m_WdApp.Selection.Range;
                        m_WdApp.ActiveDocument.Bookmarks.Add(bmkName, ref oRange);
                        break;
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private void AddAnoColNumPair(Word.Bookmark bmk, Dictionary<string, int> anoColumnDictionary)
        {
            foreach (string ano in Ano.AnoCollection)
            {
                if (ano != Ano.AnoXM)
                {
                    Word.Cell cellNC = GetAnoCell(ano, bmk);
                    if (cellNC != null)
                    {
                        anoColumnDictionary.Add(ano, cellNC.ColumnIndex);
                        break;
                    }
                }
            }
        }

        private void ToRight(Word.Cell cell)
        {
            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }

        private Word.Cell GetAnoCell(string anoStarts, Word.Bookmark bmk)
        {
            try
            {
                if (bmk.Name.ToUpper().StartsWith(anoStarts) && !bmk.Name.EndsWith("_"))
                {
                    bmk.Range.Select();
                    m_WdApp.Selection.SelectCell();
                    Word.Cell cellAno = m_WdApp.Selection.Cells.Item(1);
                    return cellAno;
                }
            }
            catch (Exception)
            {
                //throw;
            }
            return null;
        }

        #region IDisposable 成员
        public void Dispose()
        {
            try
            {
                object saveChanges = false;
                ((Word._Application)m_WdApp).Quit(ref saveChanges, ref Optional, ref Optional);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("结束COM服务器错误: " + ex.Message);
            }
            finally
            {
                OfficeWrap.NAR(m_WdApp);
            }
        }
        #endregion
    }
}
