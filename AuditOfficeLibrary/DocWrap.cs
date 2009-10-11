using System;
using System.Collections.Generic;
using System.Text;
using AuditPubLib;

namespace AuditOfficeLibrary
{
    public class DocWrap
    {
        int _id = -1;

        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }

        object _doc;

        public object Doc
        {
            get { return _doc; }
            set { _doc = value; }
        }


        int _editState; //0:New 1:Created 2:Editing

        /// <summary>
        /// 0:New 1:Created 2:Editing
        /// </summary>
        public int EditState
        {
            get { return _editState; }
            set { _editState = value; }
        }

        string _path;

        public string Path          
        {
            get { return _path; }
            set { _path = value; }
        }

        private int _projectId = -1;

        public int ProjectId
        {
            get { return _projectId; }
            set { _projectId = value; }
        }

        private int _directoryId = -1;

        public int DirectoryId
        {
            get { return _directoryId; }
        }

        public string DirectoryName
        {
            get { return DAL.GetTemplateDirectoryName(_directoryId, _docType); }
        }

        private DocType _docType;

        public DocType DocType
        {
            get { return _docType; }
            set { _docType = value; }
        }

        public DocWrap(AuditDataSetPRO.PRO_WorkSheetRow drWorksheet)
        {
            _id = drWorksheet.ID;
            _projectId = drWorksheet.ProjectID;
            _directoryId = drWorksheet.DirectoryID;
            _path = Common.GetTempDirectoryPath() + _id;
            _docType = DocType.Worksheet;   
        }

        public DocWrap(AuditDataSetTP.TP_WorkSheetRow drTemplate)
        {
            _id = drTemplate.ID;
            _directoryId = drTemplate.DirectoryID;
            _path = Common.GetTempDirectoryPath() + _id;
            _docType = DocType.Template;
        }

        public DocWrap(AuditDataSetPROHis.PRO_WorkSheet_HisRow drWorksheetHis)
        {
            _id = drWorksheetHis.ID;
            _projectId = drWorksheetHis.ProjectID;
            _directoryId = drWorksheetHis.DirectoryID;
            _path = Common.GetTempDirectoryPath() + _id;
            _docType = DocType.Worksheet;
        }

        public DocWrap()
        {
        }
    }

}
