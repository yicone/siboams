using System;
using System.Collections.Generic;
using System.Text;
using AuditPubLib;

namespace AuditOfficeLibrary
{
    public class BeforeInsertIndexEventArgs
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

        private FileType _fileType;

        public FileType FileType
        {
            get { return _fileType; }
            set { _fileType = value; }
        }

        public BeforeInsertIndexEventArgs(object address, object textToDisplay)
        {
            _address = address;
            _textToDisplay = textToDisplay;
        }
    }
}
