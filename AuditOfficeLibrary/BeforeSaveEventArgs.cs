using System;
using System.Collections.Generic;
using System.Text;

namespace AuditOfficeLibrary
{
    public class BeforeSaveEventArgs
    {
        private DocWrap _docWrap;

        public DocWrap DocWrap
        {
            get { return _docWrap; }
            set { _docWrap = value; }
        }

        public BeforeSaveEventArgs(DocWrap docWrap)
        {
            _docWrap = docWrap;
        }
    }
}
