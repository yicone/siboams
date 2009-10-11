using System;
using System.Collections.Generic;
using System.Text;
using Word;
using Excel;
using System.ComponentModel;
using System.Reflection;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections;
using AuditPubLib;
using Office;

namespace AuditOfficeLibrary
{
    public abstract class OfficeWrap
    {
        protected CommandBar _commandBarSBMMark = null;
        protected CommandBarButton btnAddMark, btnDeleteMark, btnInsertIndex, btnSaveMark, btnSaveAnno, btnInsertResult = null;
        private int m_FaceId = 113;
        private int m_Tag = 0;
        public static object Optional = Missing.Value;
        public static object False = false;

        //public abstract int ProjectId
        //{
        //    get;
        //    set;
        //}

        public abstract DocWrap DocWrap
        {
            get;
        }

        public abstract List<Mark> Marks
        {
            get;
        }

        #region Obsolete! 原附注标志
        /*        public abstract Dictionary<string, string> AnnoDictionary
        {
            get;
        }*/

        #endregion

        public abstract Dictionary<string, string> OtherMarkDictionary
        {
            get;
        }

        public abstract Dictionary<string, string> RefedMarkDictionary
        {
            get;
        }
        

        public static void NAR(object o)
        {
            try
            {
                Marshal.ReleaseComObject(o);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("释放Com对象出现异常: " + ex.Message);
            }
            finally
            {
                o = null;
            }
        }

        /// <summary>
        /// 获取指定全路径的全文件名
        /// </summary>
        /// <param name="fullName"></param>
        /// <returns></returns>
        public static string GetFullName(string fullPath)
        {
            string fullName = fullPath;

            string extension = Path.GetExtension(fullPath);

            if (extension != "")
            {
                int i = fullPath.LastIndexOf(extension);
                Debug.Assert(i != -1);
                fullName = fullPath.Remove(i);
            }

            return fullName;
        }

        public static bool DownloadDoc(Byte[] blob, string fileType, string tempDocPath)
        {
            if (blob == null) throw new NullReferenceException();
            string tempFileName = "";

            switch (fileType)
            {
                case ".doc":
                    tempFileName = tempDocPath + ".doc";
                    break;
                case ".xls":
                    tempFileName = tempDocPath + ".xls";
                    break;
                default:
                    throw new ArgumentException("Unsupport FileType!");
            }

            try
            {
                Blob.Write(blob, tempFileName);
                return true;
            }
            catch (Exception ex)
            {
                string errMsg = "下载文档时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                //throw new Exception(errMsg, ex);
            }

            return false;
        }

        /// <summary>
        /// 得到文档内容的字节数组
        /// 注意:确保子类覆盖了该方法,否则new FileStrem()得不到正确的参数
        /// </summary>
        /// <param name="blob"></param>
        public virtual Byte[] GetDocBytes(string fullPath)
        {
            Byte[] blob = null;

            try
            {
                //FileShare.ReadWrite保证了读取正在编辑的Word内容时,
                //不会出现"另一进程正在使用此文件的错误".
                using (FileStream stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    blob = new Byte[stream.Length];
                    stream.Read(blob, 0, blob.Length);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "获取文档内容时出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }

            return blob;
        }

        public virtual string GetDocText(string filePath)
        {
            return "";
        }

        public virtual List<string> GetCrossRefList()
        {
            return null;
        }

        //因为需要Work/Excel的实例,所以将此方法交由子类公开
        protected virtual void AddCommandBar()
        {
            AddCommandBarButtons();
        }

        //可以由子类增加或删除按钮,注册或取消注册按钮点击事件处理函数
        protected void AddCommandBarButtons()
        {
            try
            {
                AddCommandBarButton(ref btnAddMark,     "插入智能标志");
                AddCommandBarButton(ref btnDeleteMark,  "删除智能标志");
                AddCommandBarButton(ref btnInsertIndex, "插入交叉索引");
                AddCommandBarButton(ref btnSaveMark, "保存标志");
                AddCommandBarButton(ref btnSaveAnno, "保存附注");
                AddCommandBarButton(ref btnInsertResult, "审计说明");

                btnAddMark.Click += new _CommandBarButtonEvents_ClickEventHandler(btnAddMark_Click);
                btnDeleteMark.Click += new _CommandBarButtonEvents_ClickEventHandler(btnDeleteMark_Click);
                btnInsertIndex.Click += new _CommandBarButtonEvents_ClickEventHandler(btnInsertIndex_Click);
                btnSaveMark.Click += new _CommandBarButtonEvents_ClickEventHandler(btnSaveMark_Click);
                btnSaveAnno.Click += new _CommandBarButtonEvents_ClickEventHandler(btnSaveAnno_Click);
                btnInsertResult.Click += new _CommandBarButtonEvents_ClickEventHandler(btnInsertResult_Click);

                //_commandBarSBMMark.ShowPopup(OfficeWrap.Optional, OfficeWrap.Optional);
                _commandBarSBMMark.Visible = true;
            }
            catch (Exception ex)
            {
                string errMsg = "初始化工具条出现错误: " + ex.Message;
                Debug.WriteLine(errMsg);
                throw new Exception(errMsg, ex);
            }
        }

        private void AddCommandBarButton(ref CommandBarButton cbb, string buttonText)
        {
            cbb = (CommandBarButton)_commandBarSBMMark.Controls.Add(1, Optional, Optional, Optional, Optional);
            cbb.Style = MsoButtonStyle.msoButtonCaption;
            cbb.Caption = buttonText;
            cbb.FaceId = m_FaceId++;
            cbb.Tag = (m_Tag++).ToString();
        }

        protected virtual void OnInsertIndex(BeforeInsertIndexEventArgs e)
        {
            if (this.BeforeInsertIndex != null)
            {
                BeforeInsertIndex(this, e);
            }
        }

        public virtual void AppendText(string text)
        {
 
        }

        #region virtual CommandBarButton event Handler
        protected virtual void btnInsertIndex_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        protected virtual void btnDeleteMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        protected virtual void btnAddMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        protected virtual void btnSaveMark_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        protected virtual void btnSaveAnno_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        protected virtual void btnInsertResult_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
        } 
        #endregion

        #region virtual智能标记操作
        public virtual void UpdateMark(string mark)
        {
 
        }

        /// <summary>
        /// 在Office文档中删除指定名称的书签或名称
        /// </summary>
        /// <param name="mark"></param>
        public virtual void DeleteMark(string mark)
        {
 
        }
        #endregion

        public event BeforeInsertIndexHandler BeforeInsertIndex;
    }

    public delegate void BeforeInsertIndexHandler(object sender, BeforeInsertIndexEventArgs e);
    public delegate void DocumentBeforeSaveHandler(object sender, BeforeSaveEventArgs e);
    public delegate void DocumentAfterSaveHandler(object sender, AfterSaveEventArgs e);
}
