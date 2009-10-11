using System;
using System.Collections.Generic;
using System.Text;

namespace AuditOfficeLibrary
{
    public delegate void InsertCrossRefHandler(InsertCrossRefEventArgs e);

    public class InsertCrossRefEventArgs : EventArgs
    {
        private object _address;

        public object Address
        {
            get { return _address; }
            set { _address = value; }
        }

        private object _textToDisplay;

        public object TextToDisplay
        {
            get { return _textToDisplay; }
            set { _textToDisplay = value; }
        }

        public InsertCrossRefEventArgs(object address, object textToDisplay)
        {
            _address = address;
            _textToDisplay = textToDisplay;
        }
    }

}
