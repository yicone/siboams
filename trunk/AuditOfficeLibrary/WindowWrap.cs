using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace AuditOfficeLibrary
{
    public class WindowWrap :IWin32Window
    {
        private IntPtr m_Hwnd;

        #region IWin32Window ≥…‘±

        public IntPtr Handle
        {
            get { return m_Hwnd; }
        }

        #endregion

        public WindowWrap(IntPtr handle)
        {
            m_Hwnd = handle;
        }
    }
}
