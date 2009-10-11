using System;
using System.Collections.Generic;
using System.Text;
using Office;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Reflection;
using Word;

namespace AuditOfficeLibrary
{
    public delegate void DocumentSavedHandler(DocumentSavedEventArgs e);

    public class DocumentSavedEventArgs : EventArgs
    {
        private Document m_Doc;
        private bool m_Successed;
        private string m_FileName;

        public string FileName
        {
            get { return m_FileName; }
            set { m_FileName = value; }
        }

        public bool Successed
        {
            get { return m_Successed; }
            set { m_Successed = value; }
        }

        public Document Doc
        {
            get { return m_Doc; }
            set { m_Doc = value; }
        }

        public DocumentSavedEventArgs(Document doc, string fileName, bool successed)
        {
            m_Doc = doc;
            m_Successed = successed;
            m_FileName = fileName;
        }
    }

    public class ManagedWord
    {
        private static object c_Optional = Missing.Value;
        private static object c_OFalse = false;
        private static Word.Application s_WdApp = null;
        private static bool m_CanSave = true;
        private static bool s_BeAboutToClose = false;
        static CommandBarButton s_DocBtnInsertMark, s_DocBtnDeleteMark, s_DocBtnInsertCrossRef, s_DocBtnSaveMark, s_DocBtnSaveAnnotation, s_DocBtnInsertResult;

        public static bool CanSave
        {
            get { return ManagedWord.m_CanSave; }
            set { ManagedWord.m_CanSave = value; }
        }


        public static event DocumentSavedHandler DocumentSaved;
        public static event InsertCrossRefHandler InsertCrossRef;

        protected static void OnInsertCrossRef(InsertCrossRefEventArgs e)
        {
            if (InsertCrossRef != null)
            {
                InsertCrossRef(e);
            }
        }

        protected static void OnDocumentSaved(DocumentSavedEventArgs e)
        {
            if (DocumentSaved != null)
            {
                DocumentSaved(e);
            }
        }

        public static Word.Application GetApplication()
        {
            bool isCreateApp = false;
            List<Process> goodXlProcList = ScanWordProc(WdProcQuality.Good);
            if (s_WdApp == null || goodXlProcList.Count == 0)
            {
                isCreateApp = true;
            }
            else
            {
                try
                {
                    if (s_WdApp.Visible) ;
                }
                catch
                {
                    isCreateApp = true;
                }
            }

            //if (s_WdApp == null)
            if (isCreateApp)
            {
                s_WdApp = new Word.Application();
                //s_DocApp.DisplayAlerts = false;  //NOTE!
                s_WdApp.Visible = true;
                s_WdApp.DocumentBeforeSave += new ApplicationEvents2_DocumentBeforeSaveEventHandler(s_WdApp_DocumentBeforeSave);
                s_WdApp.DocumentBeforeClose += new ApplicationEvents2_DocumentBeforeCloseEventHandler(s_WdApp_DocumentBeforeClose);

                CommandBar docCommandBar;
                try
                {
                    docCommandBar = s_WdApp.CommandBars["智能标志"];
                    docCommandBar.Visible = true;
                }
                catch (Exception ex)
                {
                    Log.Write(ex.Message + ex.StackTrace);

                    docCommandBar = s_WdApp.CommandBars.Add("智能标志", MsoBarPosition.msoBarFloating,
                    c_Optional, true);
                    docCommandBar.Width = 100;
                    #region Add CommandBarButtons
                    s_DocBtnInsertMark = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnInsertMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnInsertMark.Caption = "插入标志";
                    s_DocBtnInsertMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnInsertMark_Click);
                    s_DocBtnDeleteMark = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnDeleteMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnDeleteMark.Caption = "删除标志";
                    s_DocBtnDeleteMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnDeleteMark_Click);
                    s_DocBtnInsertCrossRef = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnInsertCrossRef.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnInsertCrossRef.Caption = "插入交叉索引";
                    s_DocBtnInsertCrossRef.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnInsertCrossRef_Click);
                    s_DocBtnSaveMark = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnSaveMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnSaveMark.Caption = "保存标志";
                    s_DocBtnSaveMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnSaveMark_Click);
                    s_DocBtnSaveAnnotation = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnSaveAnnotation.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnSaveAnnotation.Caption = "保存附注";
                    s_DocBtnSaveAnnotation.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnSaveAnnotation_Click);
                    s_DocBtnInsertResult = (CommandBarButton)docCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_DocBtnInsertResult.Style = MsoButtonStyle.msoButtonCaption;
                    s_DocBtnInsertResult.Caption = "审计说明";
                    s_DocBtnInsertResult.Click += new _CommandBarButtonEvents_ClickEventHandler(s_DocBtnInsertResult_Click);
                    #endregion
                    docCommandBar.Visible = true;
                    ////NOTE:确保不修改到Normal.dot
                    object fileName = "Normal.dot";
                    try
                    {
                        s_WdApp.Templates.Item(ref fileName).Saved = true;
                    }
                    catch (Exception ex2)
                    {
                        Log.Write(ex2.Message + ex2.StackTrace, false);
                    }

                }
            }
            return s_WdApp;
        }

        #region CommandBarButton Click Evnet Hanlder

        static void ShowDialog(object oModalForm)
        {
            System.Windows.Forms.Form modalForm = (System.Windows.Forms.Form)oModalForm;
            modalForm.ShowDialog();
        }

        static void s_DocBtnInsertMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Word.Application wdApp = Ctrl.Application as Word.Application;
                InsertMarkForm frmInsertMark = new InsertMarkForm(wdApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frmInsertMark);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_DocBtnDeleteMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Word.Application wdApp = Ctrl.Application as Word.Application;
                DeleteMarkForm frmDeleteMark = new DeleteMarkForm(wdApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frmDeleteMark);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_DocBtnInsertCrossRef_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            object address = null;
            object textToDisplay = null;
            InsertCrossRefEventArgs e = new InsertCrossRefEventArgs(address, textToDisplay);
            OnInsertCrossRef(e);

            if (e.Address != null && e.TextToDisplay != null)
            {
                Word.Application wdApp = Ctrl.Application as Word.Application;

                try
                {
                    address = e.Address;
                    textToDisplay = e.TextToDisplay;
                    wdApp.ActiveDocument.Hyperlinks.Add(wdApp.Selection.Range,
                                    ref address, ref c_Optional, ref c_Optional, ref textToDisplay, ref c_Optional);
                    //
                    wdApp.Activate();
                }
                catch (Exception ex)
                {
                    Log.Write(ex.Message + ex.StackTrace, true);
                }
            }
        }

        static void s_DocBtnSaveMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Word.Application wdApp = Ctrl.Application as Word.Application;
                if (wdApp.Selection.Cells.Count > 0)
                {
                    wdApp.Selection.SelectCell();
                    Cell wdCell = wdApp.Selection.Cells.Item(1);
                    if (wdCell == null) return;

                    int rowIndex = wdCell.RowIndex;
                    int colIndex = wdCell.ColumnIndex;
                    string mark = "";
                    if (colIndex - 1 >= 0)
                    {
                        try
                        {
                            mark = wdCell.Previous.Range.Text.Replace("\r\a", "");
                        }
                        catch { }

                        InsertOtherMarkForm frm = new InsertOtherMarkForm(wdApp, mark);
                        Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                        thread.Start(frm);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_DocBtnSaveAnnotation_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //try
            //{
            //    Word.Application wdApp = Ctrl.Application as Word.Application;
            //    if (wdApp.Selection.Cells.Count > 0)
            //    {
            //        wdApp.Selection.SelectCell();
            //        Cell wdCell = wdApp.Selection.Cells.Item(1);
            //        if (wdCell == null) return;

            //        int rowIndex = wdCell.RowIndex;
            //        int colIndex = wdCell.ColumnIndex;
            //        if (colIndex - 1 >= 0)
            //        {
            //            InsertAnoForm frm = new InsertAnoForm(wdApp);
            //            Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
            //            thread.Start(frm);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Log.Write(ex.Message + ex.StackTrace, true);
            //}
        }

        static void s_DocBtnInsertResult_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Word.Application wdApp = Ctrl.Application as Word.Application;
                InsertResultForm frm = new InsertResultForm(wdApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frm);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }
        #endregion

        public static Document NewDocument()
        {
            s_BeAboutToClose = false;
            Word.Application wdApp = GetApplication();
            Document doc = wdApp.Documents.Add(ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional);
            wdApp.Visible = true;
            //NOTE:曾经发生无法立刻保存的问题,曾用延时300.
            return doc;
        }

        public static Document OpenDocument(string fileName)
        {
            return OpenDocument(fileName, false, false);
        }

        public static Document OpenDocument(string fileName, bool quiet, bool readOnly)
        {
            s_BeAboutToClose = false;
            Word.Application wdApp;
            object oVisable;

            if (!quiet)
            {
                wdApp = GetApplication();
                oVisable = true;
            }
            else
            {
                wdApp = new Application();
                oVisable = false;
            }

            object oFileName = fileName;
            object oReadOnly = readOnly;

            Document doc = wdApp.Documents.Open(ref oFileName, ref c_Optional, ref oReadOnly, ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, ref oVisable);

            if(!quiet)
                wdApp.Activate();

            return doc;
        }

        public static bool SaveDocument(Document newDoc, string fileName)
        {
            if (!newDoc.Name.EndsWith(".doc"))
            {
                if (!fileName.EndsWith(".doc"))
                {
                    fileName += ".doc";
                }

                try
                {
                    object oFileName = fileName;
                    newDoc.SaveAs(ref oFileName, 
                        ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, 
                        ref c_Optional, ref c_Optional, ref c_Optional, ref c_Optional, 
                        ref c_Optional, ref c_Optional);
                }
                catch (Exception ex)
                {
                    Log.Write(ex.Message + ex.StackTrace);
                    return false;
                }
                return true;
            }
            return false;
        }

        public static string GetTextContent(Document wdDoc)
        {
            string text = "";
            text = wdDoc.Content.Text;
            return text;
        }

        public static List<string> GetCrossRefList(Document wdDoc)
        {
            List<string> crossRefList = new List<string>();
            string wdName = wdDoc.Name;

            foreach (Hyperlink hl in wdDoc.Hyperlinks)
            {
                string refedFileName = hl.Address;
                if (!string.IsNullOrEmpty(refedFileName) && !crossRefList.Contains(refedFileName) && refedFileName != wdName)
                {
                    crossRefList.Add(refedFileName);
                }
            }

            return crossRefList;
        }

        public static Dictionary<string, Bookmark> GetBookmarkDict(Document wdDoc)
        {
            Dictionary<string, Bookmark> docNameDict = new Dictionary<string, Bookmark>();
            foreach (Bookmark wdBookmark in wdDoc.Bookmarks)
            {
                docNameDict.Add(wdBookmark.Name, wdBookmark);
            }

            return docNameDict;
        }

        public static void ReplaceMarks(Document wdDoc, List<Mark> marks)
        {
            Application wdApp = wdDoc.Application;
            List<Bookmark> bookmarkList = new List<Bookmark>();
            foreach (Bookmark bookmark in wdDoc.Bookmarks)
            {
                bookmarkList.Add(bookmark);
            }

            foreach (Bookmark bookmark in bookmarkList)
            {
                foreach (Mark mark in marks)
                {
                    if (mark.Formula == bookmark.Name)
                    {
                        Range range = bookmark.Range;
                        if (mark.Value.Contains("|"))
                        {
                            range.Select();
                            string[] paragraphs = mark.Value.Split('|');
                            for (int i = 0; i < paragraphs.Length; i++)
                            {
                                string str = paragraphs[i];
                                if (i != 0)
                                {
                                    wdApp.Selection.TypeParagraph();
                                }
                                wdApp.Selection.TypeText(str);
                            }

                            range.SetRange(range.Start, wdApp.Selection.End);
                        }
                        else
                        {
                            range.Text = mark.Value;
                        }
                        range.Select();
                        object oRange = wdApp.Selection.Range;
                        wdDoc.Bookmarks.Add(mark.Formula, ref oRange);
                        break;
                    }
                }
            }
        }

        public static void SpecialPaste(Document wdDoc)
        {
            object units = Word.WdUnits.wdStory;
            wdDoc.Application.Selection.EndKey(ref units, ref c_Optional);
            wdDoc.Application.Selection.Paste();
        }

        public static void CopyAll(Word.Application wdApp)
        {
            wdApp.Selection.WholeStory();
            wdApp.Selection.Copy();
        }

        public static void InsertPageBreak(Word.Document wdDoc)
        {
            Word.Application wdApp = wdDoc.Application;
            if (wdApp.Selection != null)
            {
                object oEnd = wdApp.Selection.End;
                object oBreakType = WdBreakType.wdPageBreak;
                Word.Range wdRng = wdDoc.Range(ref oEnd, ref oEnd);
                wdRng.InsertBreak(ref oBreakType);
            }
        }

        static void s_WdApp_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            s_BeAboutToClose = true;
            Log.Write("s_WdApp_DocumentBeforeClose():Doc:" + Doc.Name);
        }

        static void s_WdApp_DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //if (m_CanSave && !SaveAsUI)
            if(m_CanSave)
            {
                Log.Write("s_WdApp_DocumentBeforeSave():Doc:" + Doc.Name);
                string fileName = Path.Combine(Doc.Path, Doc.Name);
                if (!fileName.EndsWith(".doc")) return;

                Thread thread = new Thread(new ParameterizedThreadStart(AsyncOpeDocument));
                thread.Start(new AsyncOpeDocumentParam(fileName, Doc));
                Log.Write("s_WdApp_DocumentBeforeSave() end.");
            }
            else
            {  
                Cancel = true;
            }
        }

        //class
        private class AsyncOpeDocumentParam
        {
            string m_FileName;

            public string FileName
            {
                get { return m_FileName; }
                set { m_FileName = value; }
            }
            Document m_WdDoc;

            public Document Doc
            {
                get { return m_WdDoc; }
                set { m_WdDoc = value; }
            }

            public AsyncOpeDocumentParam(string fileName, Document wdDoc)
            {
                m_FileName = fileName;
                m_WdDoc = wdDoc;
            }
        }

        static void AsyncOpeDocument(object oAsyncOpeDocumentParam)
        {
            Log.Write("AsyncOpeDocument start...");

            bool saveSuccessed = true;
            AsyncOpeDocumentParam param = (AsyncOpeDocumentParam)oAsyncOpeDocumentParam;
            Document wdDoc = param.Doc;
            string fileName = param.FileName;

            IntPtr p = AuditPubLib.Common.FindWindow(null, "另存为");
            if (p != IntPtr.Zero)
            {
                Debug.WriteLine("检测到Word另存为窗体已打开,跳过保存.");
                return;
            }

            try
            {
                int waitTime = 0;
                while (!wdDoc.Saved)
                {
                    Thread.Sleep(100);
                    waitTime += 100;

                    if (waitTime > 3000)
                    {
                        saveSuccessed = false;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                saveSuccessed = false;
                Log.Write(ex.Message + ex.StackTrace);
            }
            finally
            {
                saveSuccessed = s_BeAboutToClose ? false : saveSuccessed;
                s_BeAboutToClose = false;
                OnDocumentSaved(new DocumentSavedEventArgs(wdDoc, fileName, saveSuccessed));
            }
            Log.Write("AsyncOpeDocument end");
        }

        public static List<Process> ScanWordProc(WdProcQuality procQuality)
        {
            List<Process> procList = new List<Process>();
            Process[] xlProcesses = Process.GetProcessesByName("WINWORD");

            foreach (Process xlProc in xlProcesses)
            {
                string xlMainWindowTitle = xlProc.MainWindowTitle.Trim();

                if (procQuality == WdProcQuality.All)
                {
                    procList.Add(xlProc);
                }
                else if (procQuality == WdProcQuality.Good && !String.IsNullOrEmpty(xlMainWindowTitle))
                {
                    procList.Add(xlProc);
                }
                else if (procQuality == WdProcQuality.Bad && String.IsNullOrEmpty(xlMainWindowTitle))
                {
                    procList.Add(xlProc);
                }
            }

            return procList;
        }

        public static void KillWord(Process proc)
        {
            proc.Kill();
        }
    }

    public enum WdProcQuality
    {
        Bad,
        Good,
        All
    }
}
