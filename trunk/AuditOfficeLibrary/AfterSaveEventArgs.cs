using System;
using System.Collections.Generic;
using System.Text;

namespace AuditOfficeLibrary
{
    public class AfterSaveEventArgs
    {
        private DocWrap _docWrap;

        public DocWrap DocWrap
        {
            get { return _docWrap; }
            set { _docWrap = value; }
        }

        public AfterSaveEventArgs(DocWrap docWrap)
        {
            _docWrap = docWrap;
        }
    }
}
