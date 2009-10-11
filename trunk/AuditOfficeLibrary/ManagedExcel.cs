using System;
using System.Collections.Generic;
using System.Text;
using Excel;
using Office;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Reflection;
using System.Windows.Forms;

namespace AuditOfficeLibrary
{
    public delegate void WorkbookSavedHandler(WorkBookSavedEventArgs e);

    public class WorkBookSavedEventArgs : EventArgs
    {
        private Workbook m_Wb;
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

        public Workbook Wb
        {
            get { return m_Wb; }
            set { m_Wb = value; }
        }

        public WorkBookSavedEventArgs(Workbook Wb, string fileName, bool successed)
        {
            m_Wb = Wb;
            m_Successed = successed;
            m_FileName = fileName;
        }
    }

    public class ManagedExcel
    {
        private static readonly object c_Optional = Missing.Value;
        private static readonly object c_OFalse = false;
        private static Excel.Application s_XlApp = null;
        private static bool s_CanSave = true;
        private static bool s_BeAboutToClose = false;
        static CommandBarButton s_XlBtnInsertMark, s_XlBtnDeleteMark, s_XlBtnInsertCrossRef, s_XlBtnSaveMark, s_XlBtnSaveAnnotation, s_XlBtnInsertResult;

        public static bool CanSave
        {
            get { return s_CanSave; }
            set { s_CanSave = value; }
        }


        public static event WorkbookSavedHandler WorkbookSaved;
        public static event InsertCrossRefHandler InsertCrossRef;

        protected static void OnInsertCrossRef(InsertCrossRefEventArgs e)
        {
            if (InsertCrossRef != null)
            {
                InsertCrossRef(e);
            }
        }

        protected static void OnWorkbookSaved(WorkBookSavedEventArgs e)
        {
            if (WorkbookSaved != null)
            {
                WorkbookSaved(e);
            }
        }

        public static Excel.Application GetApplication()
        {
            bool isCreateApp = false;
            List<Process> badXlProcList = ScanExcelProc(XlProcQuality.Bad);
            badXlProcList.ForEach(KillExcel);
            List<Process> goodXlProcList = ScanExcelProc(XlProcQuality.Good);

            if (s_XlApp == null || (badXlProcList.Count > 0 && goodXlProcList.Count == 0) || goodXlProcList.Count == 0)
            {
                isCreateApp = true;
            }
            else
            {
                try
                {
                    if (s_XlApp.Visible) ;
                }
                catch
                {
                    isCreateApp = true;
                }
            }

            if (isCreateApp)
            {
                s_XlApp = new Excel.Application();
                //s_XlApp.DisplayAlerts = false;  //NOTE!
                s_XlApp.Visible = true;
                s_XlApp.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(s_XlApp_WorkbookBeforeSave);
                s_XlApp.WorkbookBeforeClose += new AppEvents_WorkbookBeforeCloseEventHandler(s_XlApp_WorkbookBeforeClose);
                //test
                s_XlApp.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(s_XlApp_WorkbookOpen);

                CommandBar xlCommandBar;
                try
                {
                    xlCommandBar = s_XlApp.CommandBars["智能标志"];
                    xlCommandBar.Visible = true;
                }
                catch (Exception ex)
                {
                    Log.Write(ex.Message + ex.StackTrace);

                    xlCommandBar = s_XlApp.CommandBars.Add("智能标志", MsoBarPosition.msoBarFloating,
                    c_Optional, true);
                    xlCommandBar.Width = 100;
                    #region Add CommandBarButtons
                    s_XlBtnInsertMark = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnInsertMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnInsertMark.Caption = "插入标志";
                    s_XlBtnInsertMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnInsertMark_Click);
                    s_XlBtnDeleteMark = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnDeleteMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnDeleteMark.Caption = "删除标志";
                    s_XlBtnDeleteMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnDeleteMark_Click);
                    s_XlBtnInsertCrossRef = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnInsertCrossRef.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnInsertCrossRef.Caption = "插入交叉索引";
                    s_XlBtnInsertCrossRef.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnInsertCrossRef_Click);
                    s_XlBtnSaveMark = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnSaveMark.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnSaveMark.Caption = "保存标志";
                    s_XlBtnSaveMark.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnSaveMark_Click);
                    s_XlBtnSaveAnnotation = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnSaveAnnotation.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnSaveAnnotation.Caption = "保存附注";
                    s_XlBtnSaveAnnotation.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnSaveAnnotation_Click);
                    s_XlBtnInsertResult = (CommandBarButton)xlCommandBar.Controls.Add(1, c_Optional, c_Optional, c_Optional, c_Optional);
                    s_XlBtnInsertResult.Style = MsoButtonStyle.msoButtonCaption;
                    s_XlBtnInsertResult.Caption = "审计说明";
                    s_XlBtnInsertResult.Click += new _CommandBarButtonEvents_ClickEventHandler(s_XlBtnInsertResult_Click);
                    #endregion
                    xlCommandBar.Visible = true;
                }
            }
            return s_XlApp;
        }

        static void s_XlApp_WorkbookOpen(Workbook Wb)
        {
            //MessageBox.Show(Wb.Name);
        }

        #region CommandBarButton Click Evnet Hanlder

        static void ShowDialog(object oModalForm)
        {
            Form modalForm = (Form)oModalForm;
            modalForm.ShowDialog();
        }

        static void s_XlBtnInsertMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                InsertMarkForm frmInsertMark = new InsertMarkForm(xlApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frmInsertMark);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_XlBtnDeleteMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                DeleteMarkForm frmDeleteMark = new DeleteMarkForm(xlApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frmDeleteMark);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_XlBtnInsertCrossRef_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            object address = null;
            object textToDisplay = null;
            InsertCrossRefEventArgs e = new InsertCrossRefEventArgs(address, textToDisplay);
            OnInsertCrossRef(e);

            if (e.Address != null && e.TextToDisplay != null)
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                if (xlApp != null)
                {
                    try
                    {
                        ((Excel.Worksheet)xlApp.ActiveSheet).Hyperlinks.Add(xlApp.ActiveCell,
                                        e.Address.ToString(), c_Optional, c_Optional, e.TextToDisplay);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex.Message + ex.StackTrace, true);
                    }
                }
            }
        }

        static void s_XlBtnSaveMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                Range xlRange = xlApp.ActiveCell;
                int rowIndex = xlRange.Row;
                int colIndex = xlRange.Column;
                string mark = "";
                if (colIndex - 1 >= 0)
                {
                    try
                    {
                        mark = ((Excel.Range)xlApp.Cells[rowIndex, colIndex - 1]).Text.ToString();
                    }
                    catch { }
                }

                //string markValue = _excelSigleton.ActiveCell.Text.ToString();

                InsertOtherMarkForm frm = new InsertOtherMarkForm(xlApp, mark);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frm);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_XlBtnSaveAnnotation_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                InsertAnoForm frm = new InsertAnoForm(xlApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frm);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }

        static void s_XlBtnInsertResult_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Excel.Application xlApp = Ctrl.Application as Excel.Application;
                InsertResultForm frm = new InsertResultForm(xlApp);
                Thread thread = new Thread(new ParameterizedThreadStart(ShowDialog));
                thread.Start(frm);
            }
            catch (Exception ex)
            {
                Log.Write(ex.Message + ex.StackTrace, true);
            }
        }
        #endregion

        public static Workbook NewWorkbook()
        {
            s_BeAboutToClose = false;
            Excel.Application xlApp = GetApplication();
            Workbook xlWb = xlApp.Workbooks.Add(c_Optional);
            xlApp.Visible = true;
            return xlWb;
        }

        public static Workbook OpenWorkbook(string fileName)
        {
            return OpenWorkbook(fileName, false, false);
        }

        public static Workbook OpenWorkbook(string fileName, bool quiet, bool readOnly)
        {
            s_BeAboutToClose = false;
            Excel.Application xlApp;
            Workbook xlWb = null;
            object updateLinks = c_Optional;

            if (!quiet)
            {
                xlApp = GetApplication();
            }
            else
            {
                xlApp = new Excel.Application();
                updateLinks = 2;
            }


            xlWb = xlApp.Workbooks.Open(fileName, updateLinks, readOnly, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional);
            xlWb.Activate();

            return xlWb;
        }

        public static bool SaveWorkbook(Workbook newWb, string fileName)
        {
            if (!newWb.Name.EndsWith(".xls"))
            {
                if (!fileName.EndsWith(".xls"))
                {
                    fileName += ".xls";
                }

                newWb.SaveAs(fileName, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, XlSaveAsAccessMode.xlNoChange, c_Optional, c_Optional, c_Optional, c_Optional);
                return true;
            }
            return false;
        }

        public static string GetTextContent(Workbook wb)
        {
            string tempFileName = AuditPubLib.Common.GetTempDirectoryPath() + "temp";
            Workbook tempXlBook = null;
            Excel.Application xlApp = wb.Application;
            List<string> txtFileNameList;

            try
            {
                //目的是不显示对打开文档时提示更新对外部文件公式的计算
                xlApp.DisplayAlerts = false;        //NOTE:无效,原因?

                wb.SaveCopyAs(tempFileName);
                object isUpdateLinks = 2;   // Never update links for this workbook on opening 
                tempXlBook = xlApp.Workbooks.Open(tempFileName, isUpdateLinks, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional);
                //上一句导致wb.Saved = false,恢复之
                wb.Saved = true;
                //NOTE:false,原因?
                //Debug.Assert(object.ReferenceEquals(xlApp,tempXlBook.Application));

                txtFileNameList = new List<string>();
                string tempTxtFileName;
                foreach (Worksheet xlSheet in tempXlBook.Worksheets)
                {
                    tempTxtFileName = String.Format(@"{0}{1}.txt", tempFileName, txtFileNameList.Count);
                    xlSheet.SaveAs(tempTxtFileName, XlFileFormat.xlCSV, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional, c_Optional);
                    txtFileNameList.Add(tempTxtFileName);
                }
            }
            finally
            {
                if (tempXlBook != null)
                    tempXlBook.Close(c_OFalse, c_Optional, c_Optional);
                xlApp.DisplayAlerts = true;
            }

            StringBuilder sb = new StringBuilder();
            foreach (string txtFileName in txtFileNameList)
            {
                using (FileStream fs = new FileStream(txtFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //choice of Encodeing from a test result.
                    TextReader txtReader = new StreamReader(fs, Encoding.Default);
                    sb.Append(txtReader.ReadToEnd());
                }
            }
            return sb.ToString();
        }

        public static List<string> GetCrossRefList(Workbook wb)
        {
            List<string> crossRefList = new List<string>();
            string wbName = wb.Name;

            foreach (Excel.Worksheet xlSheet in wb.Worksheets)
            {
                foreach (Hyperlink hl in xlSheet.Hyperlinks)
                {
                    string refedFileName = hl.Address;
                    if (!crossRefList.Contains(refedFileName) && refedFileName != wbName)
                    {
                        crossRefList.Add(refedFileName);
                    }
                }

                Excel.Range xlRange = xlSheet.Cells;
                Excel.Range xlTargeRange = xlRange.Find(".xls",
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Excel.XlSearchDirection.xlNext, Missing.Value, Missing.Value);
                string formula = "";
                while (xlTargeRange != null && xlTargeRange.Formula.ToString() != formula)
                {
                    formula = xlTargeRange.Formula.ToString();
                    int start = formula.LastIndexOf('[') + 1;
                    int end = formula.LastIndexOf(".xls");
                    string id = formula.Substring(start, end - start);
                    
                    crossRefList.Add(id);

                    xlTargeRange = xlRange.FindNext(xlTargeRange);
                }
            }

            return crossRefList;
        }

        public static Dictionary<string, Name> GetNameDict(Workbook wb)
        {
            Dictionary<string, Name> xlNameDict = new Dictionary<string, Name>();
            foreach (Name xlName in wb.Names)
            {
                if (xlName.RefersTo.ToString().Contains("REF!")) continue;

                xlNameDict.Add(xlName.Name, xlName);
            }

            return xlNameDict;
        }

        public static void ReplaceMarks(Workbook wb, List<Mark> marks)
        {
            foreach (Mark mark in marks)
            {
                Excel.Range rng = (Excel.Range)((Excel.Worksheet)wb.Worksheets[mark.SheetIndex]).Cells[mark.X, mark.Y];
                string value = rng.Value2 == null ? "" : rng.Value2.ToString();
                //NOTE!
                int colonIndex = value.Contains(":") ? value.IndexOf(':') : value.IndexOf('：');
                string iniValue = value.Substring(0, colonIndex + 1);

                //保证插入时间字符串时,Excel不会将单元格的格式自动设置为时间类型,而是保持文本格式
                if (mark.Formula!= null  && (mark.Formula.Contains("JZR") || mark.Formula.Contains("RQ")))
                {
                    rng.NumberFormatLocal = "＠";
                }
                rng.Value2 = iniValue + mark.Value;
            }
        }

        public static void SpecialPaste(Workbook wb)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)wb.ActiveSheet;
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

            sheet.Paste(sheet.Cells[i, 1], null);
        }

        static void s_XlApp_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            if (Wb.Name.EndsWith(".txt")) return;
            s_BeAboutToClose = true;
            Log.Write("m_XlApp_WorkbookBeforeClose():Wb:" + Wb.Name);
        }

        static void s_XlApp_WorkbookBeforeSave(Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            if (s_CanSave && !SaveAsUI)
            {
                Log.Write("m_XlApp_WorkbookBeforeSave():Wb:" + Wb.Name);
                string fileName = Path.Combine(Wb.Path, Wb.Name);
                if (!fileName.EndsWith(".xls")) return;

                Thread thread = new Thread(new ParameterizedThreadStart(AsyncOpeWorkbook));
                thread.Start(new AsyncOpeWorkbookParam(fileName, Wb));
                Log.Write("m_XlApp_WorkbookBeforeSave() end.");
            }
            else
            {
                Cancel = true;
            }
        }

        //class
        private class AsyncOpeWorkbookParam
        {
            string m_FileName;

            public string FileName
            {
                get { return m_FileName; }
                set { m_FileName = value; }
            }
            Workbook m_Wb;

            public Workbook Wb
            {
                get { return m_Wb; }
                set { m_Wb = value; }
            }

            public AsyncOpeWorkbookParam(string fileName, Workbook wb)
            {
                m_FileName = fileName;
                m_Wb = wb;
            }
        }

        static void AsyncOpeWorkbook(object oAsyncOpeWorkbookParam)
        {
            Log.Write("AsyncOpeWorkbook start...");

            bool saveSuccessed = true;
            AsyncOpeWorkbookParam param = (AsyncOpeWorkbookParam)oAsyncOpeWorkbookParam;
            Workbook wb = param.Wb;
            string fileName = param.FileName;

            try
            {
                int waitTime = 0;
                while (!wb.Saved)
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
                OnWorkbookSaved(new WorkBookSavedEventArgs(wb, fileName, saveSuccessed));
            }
            Log.Write("AsyncOpeWorkbook end");
        }

        public static List<Process> ScanExcelProc(XlProcQuality procQuality)
        {
            List<Process> procList = new List<Process>();
            Process[] xlProcesses = Process.GetProcessesByName("Excel");

            foreach (Process xlProc in xlProcesses)
            {
                string xlMainWindowTitle = xlProc.MainWindowTitle.Trim();

                if (procQuality == XlProcQuality.All)
                {
                    procList.Add(xlProc);
                }
                else if (procQuality == XlProcQuality.Good && !String.IsNullOrEmpty(xlMainWindowTitle))
                {
                    procList.Add(xlProc);
                }
                else if (procQuality == XlProcQuality.Bad && String.IsNullOrEmpty(xlMainWindowTitle))
                {
                    procList.Add(xlProc);
                }
            }

            return procList;
        }

        public static void KillExcel(Process proc)
        {
            proc.Kill();
        }
    }

    public enum XlProcQuality
    {
        Bad,
        Good,
        All
    }
}
