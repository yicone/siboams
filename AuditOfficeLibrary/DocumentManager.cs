using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using AuditPubLib;
using System.IO;

namespace AuditOfficeLibrary
{
    public class DocumentManager<T> where T :DataRow 
    {
        private static Dictionary<string, T> s_DocumentDict = new Dictionary<string, T>();

        public static string GetDownloadedFileName(T documentDataRow)
        {
            DataRow drDocument = (DataRow)documentDataRow;
            string fileName = Common.GetTempDirectoryPath() + drDocument["ID"].ToString() + drDocument["FileType"].ToString();
            return fileName;
        }

        public static bool DownloadDocument(T documentDataRow, string fileName)
        {
            DataRow drDocument = (DataRow)documentDataRow;
            //fileName = Common.GetTempDirectoryPath() + drDocument["ID"].ToString();
            string fileType = drDocument["FileType"].ToString();
            switch (fileType)
            {
                case ".doc":
                case ".xls":
                    //fileName += fileType;
                    byte[] content = (byte[])drDocument["Content"];

                    Blob.Write(content, fileName);
                    //加入已下载文档队列
                    AddDocument(fileName, documentDataRow);
                    return true;
                    default:
                    return false;
            }
        }

        public static void AddDocument(string fileName, T documentDataRow)
        {
            if (!s_DocumentDict.ContainsKey(fileName))
                s_DocumentDict.Add(fileName, documentDataRow);
            else
                s_DocumentDict[fileName] = documentDataRow;
        }

        public static T FindDocument(string fileName)
        {
            if (s_DocumentDict.ContainsKey(fileName))
            {
                return s_DocumentDict[fileName];
            }

            return null;
        }

        //预下载所有引用的文档
        public static void PreDownloadRefedDocuments(T documentDataRow)
        {
            Type t = documentDataRow.GetType();
            DataRow drDocument = (DataRow)documentDataRow;
            if (drDocument["CrossRefID"] != null)
            {
                string crossRefId = drDocument["CrossRefID"].ToString();
                string[] refIds = crossRefId.Split(',');
                foreach (string refId in refIds)
                {
                    if (refId == "")
                    {
                        break;
                    }

                    string fileName = Common.GetTempDirectoryPath() + refId;
                    int id = Convert.ToInt32(refId);
                    DataRow dr;


                    if (t == typeof(AuditDataSetPRO.PRO_WorkSheetRow))
                        dr = DAL.FindWorksheet(id);
                    //else if (t == typeof(AuditDataSetPROHis.PRO_WorkSheet_HisRow))
                    //    dr = DAL.FindHisWorksheet(id);
                    else if (t == typeof(AuditDataSetTP.TP_WorkSheetRow))
                        dr = DAL.FindTemplate(id);
                    else
                        return;

                    byte[] content = (byte[])dr["Content"];
                    string fileType = dr["FileType"].ToString();
                    fileName += fileType;
                    try
                    {
                        Blob.Write(content, fileName);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(ex.Message + ex.StackTrace, false);
                    }
                }//end foreach
            }
        }
    }
}
