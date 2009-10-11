using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using AuditPubLib;
using AuditPubLib.AuditDataSetSBMTableAdapters;

namespace AuditOfficeLibrary
{
    public class ReplaceMark
    {
        private static string s_AccountantPolicy = "KJZC_";   //会计政策
        private static string s_TaxItem = "SX_";              //税项
        private static string s_BaseInfo = "JBQK_";           //基本情况
        private static string s_TaxReport = "SSBG_";          //税审报告(中的一个标记)
        private static string s_TaxReport1 = "SSBG1_";         //纳税调减(税审报告中)

        private static System.Data.DataTable s_markTablePRO = new System.Data.DataTable();      //项目
        private static System.Data.DataTable s_markTableWS = new System.Data.DataTable();       //底稿
        private static System.Data.DataTable s_markTableRPT = new System.Data.DataTable();      //报表
        private static System.Data.DataTable s_markTableBAL_main = new System.Data.DataTable();   //余额表主公式
        private static string[] s_BalSubMarkTags = "NC,NM,QC,QM,JF,DF,JL,DL".Split(',');
        private static string s_BalAllMonthTag = "A";           //1-12月
        private static string s_BalReplaceDirectionV = "V";
        private static string s_BalReplaceDirectionH = "H";
        private static System.Data.DataTable s_markTableBAL_sub = new System.Data.DataTable();   //余额表副公式
        private static System.Data.DataTable s_markTableOther = new System.Data.DataTable();    //其它
        private static System.Data.DataTable s_markTableAnno = new System.Data.DataTable();     //附注
        private List<Mark> m_AbalMarkList = new List<Mark>();                                   //辅助余额表
        private List<Mark> m_NbalMarkList = new List<Mark>();                                   //新增余额表
        //private List<Mark> m_TzflMarkList = new List<Mark>();

        private int m_entityId, m_projectId, m_worksheetId;       //input
        private string m_YearString;
        private List<Mark> m_ResultMarks = new List<Mark>();         //output
        //税审报告专用
        private Mark m_markTaxReport = null; 
        private Mark m_markTaxReport1 = null;
        private Mark m_markAccountantPolicy = null;
        private Mark m_markTaxItem = null;
        private Mark m_markBaseInfo = null;


        //output
        public List<Mark> MarksResult
        {
            get { return m_ResultMarks; }
        }
        //mark
        public static string BaseInfo
        {
            get { return ReplaceMark.s_BaseInfo; }
        }


        static ReplaceMark()
        {
            #region 初始化临时表
            //项目
            DataColumn colFormula = new DataColumn("Formula");
            DataColumn colX = new DataColumn("X");
            DataColumn colY = new DataColumn("Y");
            DataColumn colSheetIndex = new DataColumn("SheetIndex");
            DataColumn colMarkMean = new DataColumn("MarkMean");
            s_markTablePRO.Columns.AddRange(new DataColumn[] { colFormula, colX, colY, colSheetIndex, colMarkMean });
            //底稿
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableWS.Columns.AddRange(new DataColumn[] { colFormula, colX, colY, colSheetIndex, colMarkMean });
            //报表
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableRPT.Columns.AddRange(new DataColumn[] { colFormula, colX, colY, colSheetIndex, colMarkMean });
            //余额表
            DataColumn colA = new DataColumn("A");
            DataColumn colB = new DataColumn("B");
            DataColumn colC = new DataColumn("C");
            DataColumn colD = new DataColumn("D");
            DataColumn colE = new DataColumn("E");
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableBAL_main.Columns.AddRange(new DataColumn[] { colA, colB, colC, colD, colE, colFormula, colX, colY, colSheetIndex, colMarkMean });

            colA = new DataColumn("A");
            colB = new DataColumn("B");
            colC = new DataColumn("C");
            colD = new DataColumn("D");
            colE = new DataColumn("E");
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableBAL_sub.Columns.AddRange(new DataColumn[] { colA, colB, colC, colD, colE, colFormula, colX, colY, colSheetIndex, colMarkMean });
            //其它
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableOther.Columns.AddRange(new DataColumn[] { colFormula, colX, colY, colSheetIndex, colMarkMean });
            //附注
            colFormula = new DataColumn("Formula");
            colX = new DataColumn("X");
            colY = new DataColumn("Y");
            colSheetIndex = new DataColumn("SheetIndex");
            colMarkMean = new DataColumn("MarkMean");
            s_markTableAnno.Columns.AddRange(new DataColumn[] { colFormula, colX, colY, colSheetIndex, colMarkMean });
            #endregion
        }

        public ReplaceMark(List<Mark> marks, int worksheetId, int projectId, int entityId, string strYear)
        {
            s_markTablePRO.Clear();
            s_markTableWS.Clear();
            s_markTableRPT.Clear();
            s_markTableBAL_sub.Clear();
            s_markTableBAL_main.Clear();
            s_markTableOther.Clear();
            s_markTableAnno.Clear();

            if (marks.Count == 0) return;

            m_worksheetId = worksheetId;
            m_projectId = projectId;
            m_entityId = entityId;
            m_YearString = strYear;

            #region 归类标记
            foreach (Mark mark in marks)
            {
                AuditDataSetSBM.SBM_MarkRow dr = DAL.GetMarkRow(mark.Formula);
                if (dr != null)
                {
                    //mark.Formula = drMark.Mark;
                    mark.MarkMean = dr.MarkMean;
                    mark.Type = dr.Type;
                    switch (mark.Type)
                    {
                        case "项目":
                            AddPROMark(mark);
                            break;
                        case "底稿":
                            AddWSMark(mark);
                            break;
                        case "报表":
                            AddRPTMark(mark);
                            break;
                        case "余额表":
                            AddBALMark(mark);
                            break;
                        case "辅助余额表":
                            m_AbalMarkList.Add(mark);
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    string upperFormula = mark.Formula.ToUpper();

                    if (upperFormula.StartsWith("BAL"))
                    {
                        mark.Type = "余额表";
                        AddBALMark(mark);
                    }
                    else if(upperFormula.StartsWith("ABAL"))
                    {
                        mark.Type = "辅助余额表";
                        m_AbalMarkList.Add(mark);
                    }
                    else if (upperFormula.StartsWith("RPT"))
                    {
                        mark.Type = "报表";
                        AddRPTMark(mark);
                    }
                    else if (upperFormula.StartsWith("ANNO"))
                    //else if (upperFormula.StartsWith("ANO"))
                    {
                        if (!mark.Formula.EndsWith("_"))
                        {
                            mark.Type = "附注";
                            AddAnnoMark(mark);
                        }
                    }
                    else if (mark.Formula[0] < 65 || mark.Formula[0] > 90)
                    {
                        //todo:应该检查是否存在?不存在,则设置Available为false
                        if (!mark.Formula.EndsWith("_"))
                        {
                            mark.Type = "其它";
                            AddOtherMark(mark);
                        }
                    }
                    else if (upperFormula == s_TaxReport)
                    {
                        //纳税调增
                        mark.Type = "税审报告";
                        m_markTaxReport = mark;
                    }
                    else if (upperFormula == s_TaxReport1)
                    {
                        //纳税调减
                        mark.Type = "税审报告";
                        m_markTaxReport1 = mark;
                    }
                    else if (upperFormula == s_AccountantPolicy)
                    {
                        mark.Type = "会计政策";
                        m_markAccountantPolicy = mark;
                    }
                    else if (upperFormula == s_TaxItem)
                    {
                        mark.Type = "税项";
                        m_markTaxItem = mark;
                    }
                    else if (upperFormula == s_BaseInfo)
                    {
                        mark.Type = "基本情况";
                        m_markBaseInfo = mark;
                    }
                    else
                    {
                        mark.Available = false;
                    }
                }
            } 
            #endregion
        }

        public void ProcessAllMark()
        {
            ProcessProjectMark();
            ProcessWorksheetMark();
            ProcessBalMark();
            ProcessNbalMark();
            ProcessAbalMark();
            ProcessRptMark();
            ProcessOtherMark();
#if DEBUG
            foreach (Mark mark in m_ResultMarks)
            {
                Debug.WriteLine(String.Format(@"第{0}行第{1}列的值为{2}", mark.X, mark.Y, mark.Value));
            }
#endif
        }

        #region 添加标记到各分类表
        private void AddPROMark(Mark mark)
        {
            DataRow dr = s_markTablePRO.NewRow();
            dr["Formula"] = mark.Formula;
            dr["X"] = mark.X;
            dr["Y"] = mark.Y;
            dr["SheetIndex"] = mark.SheetIndex;
            dr["MarkMean"] = mark.MarkMean;
            s_markTablePRO.Rows.Add(dr);
        }
        private void AddWSMark(Mark mark)
        {
            DataRow dr = s_markTableWS.NewRow();
            dr["Formula"] = mark.Formula;
            dr["X"] = mark.X;
            dr["Y"] = mark.Y;
            dr["SheetIndex"] = mark.SheetIndex;
            dr["MarkMean"] = mark.MarkMean;
            s_markTableWS.Rows.Add(dr);
        }
        private void AddRPTMark(Mark mark)
        {
            DataRow dr = s_markTableRPT.NewRow();
            dr["Formula"] = mark.Formula;
            dr["X"] = mark.X;
            dr["Y"] = mark.Y;
            dr["SheetIndex"] = mark.SheetIndex;
            dr["MarkMean"] = mark.MarkMean;
            s_markTableRPT.Rows.Add(dr);
        }
        private void AddBALMark(Mark mark)
        {
            string[] parts = mark.Formula.Split('_');
            if (parts.Length >= 6)
            {
                m_NbalMarkList.Add(mark);
            }
            else
            {
                if (CheckIsBalSubMark(parts[1]))
                {
                    //"副"标记
                    DataRow dr = s_markTableBAL_sub.NewRow();
                    dr["A"] = parts[0];
                    dr["B"] = parts[1];
                    dr["C"] = parts[2];
                    if (parts.Length > 3)
                        dr["D"] = parts[3];
                    if (parts.Length > 4)
                        dr["E"] = parts[4];
                    dr["Formula"] = mark.Formula;
                    dr["X"] = mark.X;
                    dr["Y"] = mark.Y;
                    dr["SheetIndex"] = mark.SheetIndex;
                    dr["MarkMean"] = mark.MarkMean;
                    s_markTableBAL_sub.Rows.Add(dr);
                }
                else
                {
                    //"主"标记
                    DataRow dr = s_markTableBAL_main.NewRow();
                    dr["A"] = parts[0];
                    dr["B"] = parts[1];
                    dr["C"] = parts[2];
                    if (parts.Length > 3)
                        dr["D"] = parts[3];
                    if (parts.Length > 4)
                        dr["E"] = parts[4];
                    dr["Formula"] = mark.Formula;
                    dr["X"] = mark.X;
                    dr["Y"] = mark.Y;
                    dr["SheetIndex"] = mark.SheetIndex;
                    dr["MarkMean"] = mark.MarkMean;
                    s_markTableBAL_main.Rows.Add(dr);
                }
            }
        }
        private void AddOtherMark(Mark mark)
        {
            DataRow dr = s_markTableOther.NewRow();
            dr["Formula"] = mark.Formula;
            dr["X"] = mark.X;
            dr["Y"] = mark.Y;
            dr["SheetIndex"] = mark.SheetIndex;
            dr["MarkMean"] = mark.MarkMean;
            s_markTableOther.Rows.Add(dr);
        }
        private void AddAnnoMark(Mark mark)
        {
            DataRow dr = s_markTableAnno.NewRow();
            dr["Formula"] = mark.Formula;
            dr["X"] = mark.X;
            dr["Y"] = mark.Y;
            dr["SheetIndex"] = mark.SheetIndex;
            dr["MarkMean"] = mark.MarkMean;
            s_markTableAnno.Rows.Add(dr);
        } 
        #endregion

        #region 处理各分类表的标记,为其取值

        private void ProcessProjectMark()
        {
            if (s_markTablePRO.Rows.Count > 0)
            {
                Dictionary<string, string> result = new Dictionary<string, string>();
                //"审计单位"
                result.Add("SJDW", DAL.GetEntityName());

                string strConn = Common.DBConnString;
                using (SqlConnection conn = new SqlConnection(strConn))
                {
                    string sql = String.Format(@"
                        SELECT 
		                        E.EntityName as BSDW,
		                        Case when P.AccBeginYear is not null then
			                        case when P.AccEndyear is not null then 
				                        case when P.AccBeginYear=P.AccEndyear then
					                        Rtrim(ltrim(str(P.AccbeginYear)))
				                        else
					                        Rtrim(ltrim(str(P.AccbeginYear))) +'-'+rtrim(ltrim(str(P.Accendyear)))
				                        end
			                        else 
				                        rtrim(ltrim(str(P.accbeginyear))) end
		                        else
			                        case when p.accendyear is not null then rtrim(ltrim(str(p.accendyear))) else '' end
		                        end as KJQJ,
		                        P.DeadLine as JZR,
		                        P.PlanBeginDate as XMJHKSRQ,
		                        P.PlanEnddate as XMJHJSRQ,
		                        P.ActBeginDate as XMSJKSRQ,
		                        P.actEnddate as XMSJJSRQ,
		                        P.ManagerID as SJZZ,
		                        P.Auditor as SJZY,
		                        P.ReportNum as BGH
                        FROM PRO_PROJECT P left join sbm_Entity E on P.EntityID=E.ID 
                        WHERE P.ID={0}", m_projectId);
                    //read all project mark info.
                    SqlCommand cmmd = new SqlCommand(sql, conn);
                    conn.Open();
                    using (SqlDataReader sdr = cmmd.ExecuteReader())
                    {
                        if (sdr.Read())
                        {
                            for (int i = 0; i < sdr.FieldCount; i++)
                            {
                                string columnName = sdr.GetName(i);
                                string value;
                                //处理时间类型的公式
                                if (columnName == "JZR" || columnName == "XMJHKSRQ" || columnName == "XMJHJSRQ" || columnName == "XMSJKSRQ" || columnName == "XMSJJSRQ")
                                {
                                    if (!sdr.IsDBNull(i))
                                    {
                                        value = sdr.GetDateTime(i).ToShortDateString();
                                    }
                                    else
                                    {
                                        value = "";
                                    }
                                }
                                else
                                    value = sdr.GetValue(i).ToString();
                                result.Add(columnName, value);
                            }
                        }
                    }//end using sda

                    //替换ManagerID和Auditor(s)
                    #region 处理需要将ID转换为UserName的字段
                    SqlCommand cmmd1 = new SqlCommand(@"SELECT UserName FROM SBM_User WHERE ID=@UserID", conn);
                    SqlParameter param = new SqlParameter("@UserID", SqlDbType.Int);
                    cmmd1.Parameters.Add(param);

                    param.Value = result["SJZZ"];
                    result["SJZZ"] = cmmd1.ExecuteScalar() == null ? "" : cmmd1.ExecuteScalar().ToString();

                    string userNamesJoined = "";
                    if (result["SJZY"] != "")
                    {
                        string[] array = result["SJZY"].Split(',');
                        foreach (string str in array)
                        {
                            param.Value = Int32.Parse(str);
                            userNamesJoined += cmmd1.ExecuteScalar() == null ? "" : cmmd1.ExecuteScalar().ToString() + ",";
                        }

                        if (userNamesJoined.EndsWith(","))
                        {
                            userNamesJoined = userNamesJoined.Remove(userNamesJoined.Length - 1, 1);
                        }
                    }
                    result["SJZY"] = userNamesJoined;
                    #endregion
                }//end using conn

                foreach (DataRow rowMark_cell in s_markTablePRO.Rows)
                {
                    string formula = rowMark_cell["Formula"].ToString();
                    Mark mark = new Mark(formula, Convert.ToInt32(rowMark_cell["X"]), Convert.ToInt32(rowMark_cell["Y"]), Convert.ToInt32(rowMark_cell["SheetIndex"]));
                    mark.MarkMean = rowMark_cell["MarkMean"].ToString();

                    //处理形如"KJQJ_JZR_SJZY"这样的标记
                    if (mark.MarkMean.Contains("+"))
                    {
                        string[] array = formula.Split('_');
                        string[] array1 = mark.MarkMean.Split('+');
                        for (int i = 0; i < array.Length; i++)
                        {
                            mark.Value += array1[i] + ":";
                            mark.Value += result[array[i]] + "  ";
                        }
                    }
                    else
                    {
                        string value;
                        if (result.TryGetValue(formula, out value))
                        {
                            mark.Value = result[formula];
                        }
                        else
                        {
                            //只处理BGH1-BGH9的情况,结尾的数字如是双位,不予处理.
                            formula = formula.Remove(formula.Length - 1, 1);
                            result.TryGetValue(formula, out value);
                        }
                    }
                    m_ResultMarks.Add(mark);
                }
            }//end if
        }

        public void ProcessWorksheetMark()
        {
            if (s_markTableWS.Rows.Count > 0)
            {
                Dictionary<string, string> result = new Dictionary<string, string>();
                string strConn = Common.DBConnString;
                using (SqlConnection conn = new SqlConnection(strConn))
                {
                    string sql = String.Format(@"SELECT Auditor AS SJRY, CreateDate AS SJRQ, Reviewer AS FHRY, ReviewDate AS FHRQ, ReviewIdea AS FHYJ FROM pro_worksheet WHERE ID = {0}", m_worksheetId);

                    //read all project mark info.
                    SqlCommand cmmd = new SqlCommand(sql, conn);
                    conn.Open();
                    using (SqlDataReader sdr = cmmd.ExecuteReader())
                    {
                        if (sdr.Read())
                        {
                            for (int i = 0; i < sdr.FieldCount; i++)
                            {
                                string columnName = sdr.GetName(i);
                                string value;
                                if (columnName == "FHRQ" || columnName == "SJRQ")
                                {
                                    if (!sdr.IsDBNull(i))
                                    {
                                        value = sdr.GetDateTime(i).ToShortDateString();
                                    }
                                    else
                                    {
                                        value = "";
                                    }
                                }
                                else
                                    value = sdr.GetValue(i).ToString();
                                result.Add(columnName, value);
                            }
                        }
                    }

                    #region 处理需要将ID转换为UserName的字段
                    SqlCommand cmmd1 = new SqlCommand(@"SELECT UserName FROM SBM_User WHERE ID=@UserID", conn);
                    SqlParameter param = new SqlParameter("@UserID", SqlDbType.Int);
                    cmmd1.Parameters.Add(param);

                    param.Value = result["SJRY"];
                    result["SJRY"] = cmmd1.ExecuteScalar() == null ? "" : cmmd1.ExecuteScalar().ToString();
                    param.Value = result["FHRY"];
                    result["FHRY"] = cmmd1.ExecuteScalar() == null ? "" : cmmd1.ExecuteScalar().ToString();
                    #endregion
                }//end using

                foreach (DataRow rowMark_cell in s_markTableWS.Rows)
                {
                    string formula = rowMark_cell["Formula"].ToString();
                    Mark mark = new Mark(formula, Convert.ToInt32(rowMark_cell["X"]), Convert.ToInt32(rowMark_cell["Y"]), Convert.ToInt32(rowMark_cell["SheetIndex"]));

                    //处理形如"KJQJ_JZR_SJZY"这样的标记
                    mark.MarkMean = rowMark_cell["MarkMean"].ToString();
                    if (mark.MarkMean.Contains("+"))
                    {
                        string[] array = formula.Split('_');
                        string[] array1 = mark.MarkMean.Split('+');
                        for (int i = 0; i < array.Length; i++)
                        {
                            mark.Value += array1[i] + ":";
                            mark.Value += result[array[i]] + "  ";
                        }
                    }
                    else
                    {
                        mark.Value = result[formula];
                    }

                    m_ResultMarks.Add(mark);
                }
            }//end if
        }

        private void ProcessRptMark()
        {
            foreach (DataRow rowMark_cell in s_markTableRPT.Rows)
            {
                Mark mark = new Mark(rowMark_cell["Formula"].ToString(), Convert.ToInt32(rowMark_cell["X"]), Convert.ToInt32(rowMark_cell["Y"]), Convert.ToInt32(rowMark_cell["SheetIndex"]));
                mark.Value = Pub_Function.GetRptValFromMark(m_entityId.ToString(), m_YearString, rowMark_cell["Formula"].ToString());

                if (!String.IsNullOrEmpty(mark.Value) && double.Parse(mark.Value) != 0)
                    m_ResultMarks.Add(mark);
            }
        }

        private void ProcessBalMark()
        {
            string entityId = m_entityId.ToString();
            System.Data.DataTable dtAcc;
            bool isHor = true;        //默认横向替换

            foreach (DataRow drMainMark in s_markTableBAL_main.Rows)
            {
                int grade = -1;

                string mainMarkB = drMainMark["B"].ToString();
                string mainMarkC = drMainMark["C"].ToString();
                string mainMarkFormula = drMainMark["Formula"].ToString();
                int mainValueMarkX = Convert.ToInt32(drMainMark["X"]);
                int mainValueMarkY = Convert.ToInt32(drMainMark["Y"]);
                int mainValueSheetIndex = Convert.ToInt32(drMainMark["SheetIndex"]);

                //取科目级次,只取一位
                if (!Int32.TryParse(mainMarkC.Substring(0, 1), out grade)) continue;

                //dtAcc中, 第0列是科目代码,第1列是科目名称
                dtAcc = Pub_Function.GetMarkAcc(entityId, m_YearString, mainMarkB, grade).Tables[0];

                #region MyRegion
                foreach (DataRow drAcc in dtAcc.Rows)
                {
                    bool haveRowValue = false;
                    string subjectCode = drAcc[0].ToString();
                    string subjectName = drAcc[1].ToString();   //科目名称

                    Mark mainValueMark = new Mark(mainMarkFormula, mainValueMarkX, mainValueMarkY, mainValueSheetIndex);

                    //如果级次后附加"N",则在Excel中不显示各科目
                    if (!mainMarkC.EndsWith("N"))
                    {
                        mainValueMark.Value = subjectName;
                        m_ResultMarks.Add(mainValueMark);
                    }

                    if (s_markTableBAL_sub.Rows.Count == 0) continue;

                    //开始处理列标记
                    DataRow drSubMarkTest = s_markTableBAL_sub.Rows[0];
                    string partMonthTest = drSubMarkTest["D"].ToString();
                    string partDirectionTest = drSubMarkTest["E"].ToString();

                    //如果列标记中包含月份段,并且替换方向指明为垂直替换
                    if (!String.IsNullOrEmpty(partMonthTest) && !String.IsNullOrEmpty(partDirectionTest))
                    {
                        if (partDirectionTest == s_BalReplaceDirectionV)
                            isHor = false;
                        else if (partDirectionTest == s_BalReplaceDirectionH)
                            isHor = true;
                        else
                            continue;
                    }

                    Mark subValueMark;

                    foreach (DataRow drSubMark in s_markTableBAL_sub.Rows)
                    {
                        int subValueMarkSheetIndex = Convert.ToInt32(drSubMark["SheetIndex"]);
                        if (subValueMarkSheetIndex != mainValueSheetIndex) continue;

                        int subValueMarkX, subValueMarkY;
                        string colMarkFormula = drSubMark["Formula"].ToString();
                        string partMonth = drSubMarkTest["D"].ToString();
                        string partDirection = drSubMarkTest["E"].ToString();

                        if (isHor)
                        {
                            subValueMarkX = mainValueMarkX;
                            subValueMarkY = Convert.ToInt32(drSubMark["Y"]);
                        }
                        else
                        {
                            subValueMarkX = Convert.ToInt32(drSubMark["X"]);
                            subValueMarkY = mainValueMarkY;
                        }

                        subValueMark = new Mark(colMarkFormula, subValueMarkX, subValueMarkY, subValueMarkSheetIndex);

                        if (!String.IsNullOrEmpty(partMonth))
                        {
                            #region 包含partMonth的情况
                            bool isAllMonth = (partMonth == s_BalAllMonthTag);
                            int month = -1;

                            if (!isAllMonth)
                            {
                                int.TryParse(partMonth, out month);
                            }

                            List<string> valueList = Pub_Function.GetMonthBalValFromMark(entityId, m_YearString, subjectCode, colMarkFormula);

                            if (!CheckIsAllZero(valueList))
                            {
                                haveRowValue = true;
                                int monthMarkX, monthMarkY;
                                monthMarkX = subValueMarkX;
                                monthMarkY = subValueMarkY;

                                //if (month != -1)
                                //{
                                //    if (hor)
                                //        monthMarkY += month - 1;
                                //    else
                                //        monthMarkX += month - 1;
                                //}

                                foreach (string monthValue in valueList)
                                {
                                    Mark monthMark = new Mark(colMarkFormula, monthMarkX, monthMarkY, mainValueSheetIndex);
                                    monthMark.Value = monthValue;
                                    m_ResultMarks.Add(monthMark);

                                    if (isAllMonth)
                                    {
                                        if (isHor)
                                            monthMarkY++;
                                        else
                                            monthMarkX++;
                                    }
                                }
                            }
                            #endregion
                        }
                        else
                        {
                            #region 不包含partMonth的情况
                            subValueMark.Value = Pub_Function.GetBalValFromMark(entityId, m_YearString, subjectCode, colMarkFormula);

                            if (!haveRowValue && double.Parse(subValueMark.Value) != 0)
                            {
                                haveRowValue = true;
                            }

                            if (haveRowValue)
                                m_ResultMarks.Add(subValueMark);
                            #endregion
                        }

                        //if (hor)
                        //    colMarkY++;
                    }//end foreach 列标记

                    if (haveRowValue)
                    {
                        if (isHor)
                            mainValueMarkX++;
                        else
                            mainValueMarkY++;
                    }
                    else
                        m_ResultMarks.Remove(mainValueMark);
                }//end foreach 按行标记取出的科目名称 
                #endregion
            }//end foreach 行标记
        }

        private void ProcessNbalMark()
        {
            string entityId = m_entityId.ToString();
            foreach (Mark mark in m_NbalMarkList)
            {
                string[] parts = mark.Formula.Split('_');
                if (parts.Length > 7) continue;

                bool is7Parts = (parts.Length == 7);
                bool isHor = !(is7Parts && (parts[6] == s_BalReplaceDirectionV));

                int grade;
                string xlNum;
                DataTable dtAcc = Pub_Function.ComGetBalAcc(entityId, m_YearString, mark.Formula, out grade, out xlNum);
                int i2 = isHor ? mark.X : mark.Y;
                int iXlNum;
                if(!int.TryParse(xlNum, out iXlNum)) continue;

                foreach (DataRow drAcc in dtAcc.Rows)
                {
                    i2++;
                    string accCode = drAcc[0].ToString();

                    if (grade != -1)
                    {
                        string accName = drAcc[1].ToString();   //科目名称
                        Mark mainValueMark = isHor ? (new Mark(null, i2, iXlNum, mark.SheetIndex)) :
                            (new Mark(null, iXlNum, i2, mark.SheetIndex));
                        mainValueMark.Value = accName;
                        m_ResultMarks.Add(mainValueMark);
                    }
                    else
                    {
                        i2 = i2 - 1;
                    }

                    List<string> valueList = Pub_Function.ComGetBalVal(entityId, m_YearString, accCode, mark.Formula);

                    Mark subValueMark;
                    if (!is7Parts)
                    {
                        subValueMark = new Mark(null, i2, mark.Y, mark.SheetIndex);
                        Debug.Assert(valueList.Count == 1);
                        subValueMark.Value = valueList[0];
                        m_ResultMarks.Add(subValueMark);
                    }
                    else
                    {
                        int i3 = !isHor ? mark.X : mark.Y;
                        foreach (string value in valueList)
                        {
                            subValueMark = isHor ? (new Mark(null, i2, i3, mark.SheetIndex)) :
                                (new Mark(null, i3, i2, mark.SheetIndex));
                            subValueMark.Value = value;
                            m_ResultMarks.Add(subValueMark);
                            i3++;
                        }
                    }//end if
                }
            }
        }

        private void ProcessAbalMark()
        {
            List<Mark> mainMarkList = new List<Mark>();
            List<Mark> subMarkList = new List<Mark>();
            string[] parts;


            #region 分开主标记和副标记
            foreach (Mark mark in m_AbalMarkList)
            {
                parts = mark.Formula.Split('_');
                //todo:写入error log.
                if (parts.Length < 3) continue;
                if (CheckIsBalSubMark(parts[1]))
                {
                    subMarkList.Add(mark);
                }
                else
                {
                    mainMarkList.Add(mark);
                }
            } 
            #endregion

            string strEntityId = m_entityId.ToString();
            System.Data.DataTable dtAss;
            foreach (Mark mainMark in mainMarkList)
            {
                parts = mainMark.Formula.Split('_');
                int grade = -1;
                if(!int.TryParse(parts[2].Substring(0, 1), out grade)) continue;
                //dtAcc中,第0列是AssCode,第1列是AssName,第2列是AssType
                dtAss = Pub_Function.GetAssCode(strEntityId, m_YearString, parts[1], grade).Tables[0];
                int mainValueMarkX = mainMark.X;
                int mainValueMarkY = mainMark.Y;

                foreach (DataRow drAss in dtAss.Rows)
                {
                    Mark mainValueMark = new Mark(null, mainValueMarkX, mainValueMarkY, mainMark.SheetIndex);
                    string subjectCode = drAss[0].ToString();
                    string subjectName = drAss[1].ToString();
                    string subjectType = drAss[2].ToString();
                    bool hasSubValue = false;

                    if (!parts[2].EndsWith("N"))
                    {
                        mainValueMark.Value = subjectName;
                        //mainValueMark.Type = "辅助余额表";         //16:23 2007-3-8 USAGE:DEBUG
                        m_ResultMarks.Add(mainValueMark);
                    }

                    foreach (Mark subMark in subMarkList)
                    {
                        if (subMark.SheetIndex != mainMark.SheetIndex) continue;

                        int subValueMarkX = mainValueMarkX;
                        int subValueMarkY = subMark.Y;

                        Mark subValueMark = new Mark(null, subValueMarkX, subValueMarkY, subMark.SheetIndex);
                        string strValue = Pub_Function.GetAssBalValFromMark(strEntityId, m_YearString, subjectCode, subjectType, subMark.Formula);
                        double value;
                        if (double.TryParse(strValue, out value) && value != 0)
                        {
                            subValueMark.Value = strValue;
                            //subValueMark.Type = "辅助余额表";  //16:23 2007-3-8 USAGE:DEBUG
                            m_ResultMarks.Add(subValueMark);
                            hasSubValue = true;
                        }
                    }//end foreach

                    if (hasSubValue)
                    {
                        mainValueMarkX++;
                    }
                    else
                    {
                        m_ResultMarks.Remove(mainValueMark);
                    }
                }//end foreach
            }//end foreach
        }

        private void ProcessOtherMark()
        {
            foreach (DataRow rowMark_cell in s_markTableOther.Rows)
            {
                string formula = rowMark_cell["Formula"].ToString();

                Mark mark = new Mark(formula, Convert.ToInt32(rowMark_cell["X"]), Convert.ToInt32(rowMark_cell["Y"]), Convert.ToInt32(rowMark_cell["SheetIndex"]));

                if (formula.Contains("_"))
                {
                    string[] array = formula.Split('_');

                    double value;
                    double totalValue = 0.0;
                    for (int i = 0; i < array.Length; i++)
                    {
                        string markValue = DAL.GetOtherMarkValue(array[i], m_projectId);
                        if (Double.TryParse(markValue, out value))
                        {
                            totalValue += value;
                        }
                    }
                    mark.Value = totalValue.ToString();
                }
                else
                {
                    mark.Value = DAL.GetOtherMarkValue(formula, m_projectId);
                }

                m_ResultMarks.Add(mark);
            }
        }

        public void ProcessEntitySetupMark(string accountantPolicy, string taxItem, string baseInfo, bool useBaseInfoWorksheet)
        {
            if (m_markAccountantPolicy != null)
            {
                m_markAccountantPolicy.Value = accountantPolicy;
                m_ResultMarks.Add(m_markAccountantPolicy);
            }

            if (m_markTaxItem != null)
            {
                m_markTaxItem.Value = taxItem;
                m_ResultMarks.Add(m_markTaxItem);   
            }

            if (m_markBaseInfo != null)
            {
                if (!useBaseInfoWorksheet)
                {
                    m_markBaseInfo.Value = baseInfo;
                    m_ResultMarks.Add(m_markBaseInfo);
                }
            }

            //foreach (DataRow rowMark_cell in s_markTableAnno.Rows)
            //{
            //    string formula = rowMark_cell["Formula"].ToString();

            //    Mark mark = new Mark(formula, Convert.ToInt32(rowMark_cell["X"]), Convert.ToInt32(rowMark_cell["Y"]), Convert.ToInt32(rowMark_cell["SheetIndex"]));

            //    //if (formula.Contains("_"))
            //    //{
            //    //    string[] array = formula.Split('_');

            //    //    double value;
            //    //    double totalValue = 0.0;
            //    //    for (int i = 0; i < array.Length; i++)
            //    //    {
            //    //        string markValue = DAL.GetAnnoMarkValue(array[i], _projectId);
            //    //        if (Double.TryParse(markValue, out value))
            //    //        {
            //    //            totalValue += value;
            //    //        }
            //    //    }
            //    //    mark.Value = totalValue.ToString();
            //    //}
            //    //else
            //    //{
            //    mark.Value = DAL.GetAnnoValue(formula, m_projectId);
            //    //}

            //    m_marksResult.Add(mark);
            //}
        }

        //public void GenerateTaxAuditReport(string accountantPolicy, string taxItem, string baseInfo, bool useBaseInfoWorksheet)
        public void ProcessTaxAuditReportMark()
        {
            #region 生成纳税调增段落
            if (m_markTaxReport != null)
            {
                DataSet ds = new DataSet();
                ds.ReadXml(System.Windows.Forms.Application.StartupPath + @"\Config\TaxReportMark.xml");
                //DataColumn[] keys = new DataColumn[2];
                DataTable dt = ds.Tables[0];

                //DataColumn dcName = dt.Columns["Name"];
                //keys[0] = dcName;
                //DataColumn dcOrder = dt.Columns["Order"];
                //keys[1] = dcOrder;
                //dt.PrimaryKey = keys;

                DataView dv = new DataView(dt);
                //dv.Sort = "Number,Order ASC";
                StringBuilder sb = new StringBuilder();

                //string end = "";

                foreach (DataRowView drv in dv)
                {
                    //string markName = drv["Name"].ToString();
                    //string markValue = DAL.GetOtherMarkValue(markName, m_projectId);
                    //string markTitleBegin = drv["TitleBegin"].ToString();
                    //string formula = drv["Formula"].ToString();
                    //string markTitleEnd = drv["TitleEnd"].ToString();
                    //int number = Convert.ToInt32(drv["Number"]);        //NOTE:暂时未用
                    //int order = Convert.ToInt32(drv["Order"]);
                    //bool bEnd = Convert.ToInt32(drv["BEnd"]) == 0 ? false : true;

                    string title = drv["Title"].ToString();
                    string s1 = drv["Statement1"].ToString();
                    string s2 = drv["Statement2"].ToString();
                    string s3 = drv["Statement3"].ToString();
                    string r1 = drv["Result1"].ToString();
                    string r2 = drv["Result2"].ToString();

                    ////如果每个段落的首项没有值，则不输出此段落
                    //if (order == 0 && String.IsNullOrEmpty(markValue))
                    //{
                    //    continue;
                    //}
                    //if (!String.IsNullOrEmpty(markValue))
                    //{
                    //    if (order == 0)
                    //    {
                    //        sb.Append(String.Format(@"{0}{1}{2}", markTitleBegin, markValue, markTitleEnd));

                    //        string[] array = ParseTaxReportFormula(formula, dt);
                    //        sb.Append(array[1]);
                    //        //sb.Append(String.Format(@"|{4})、{0}={1}={2}={3}", new object[] { markName, formula, array[0], markValue, order + 1 }));
                    //        end = String.Format(@"{0}={1}={2}={3}", new object[] { markName, formula, array[0], markValue});
                    //    }

                    //    if (bEnd)
                    //    {
                    //        sb.Append(string.Format(@"|{0})、{1}", order, end));
                    //        sb.Append("|");
                    //    }
                    //}
                    sb.Append(title);
                    double totalValue = 0;
                    string f1 = ReplaceFormula(s1, ref totalValue);//1

                    if (totalValue == 0)
                    {
                        sb.Append('|');
                        sb.Append(r1);
                    }
                    else
                    {
                        sb.Append('|');
                        sb.Append(f1);
                        totalValue = 0;
                        string f2 = ReplaceFormula(s2, ref totalValue);//2
                        sb.Append('|');
                        sb.Append(f2);
                        totalValue = 0;
                        string f3 = ReplaceFormula(s3, ref totalValue);//3
                        if (totalValue > 0)
                        {
                            sb.Append('|');
                            sb.Append(f3);
                        }
                        else
                        {
                            sb.Append('|');
                            sb.Append(r2);
                        }
                    }

                    sb.Append('|');
                }

                string sbString = sb.ToString();
                string ssbgMarkValue = sbString.EndsWith("|") ? sbString.Remove(sbString.Length - 1) : sbString;
                m_markTaxReport.Value = ssbgMarkValue;
                m_ResultMarks.Add(m_markTaxReport);
            } 
            #endregion

            #region 生成纳税调减段落
            if (m_markTaxReport1 != null)
            {
                bool isAllZero = true;
                DataSet ds = new DataSet();
                ds.ReadXml(System.Windows.Forms.Application.StartupPath + @"\Config\TaxReportMark1.xml");
                DataTable dt = ds.Tables[0];
                DataView dv = new DataView(dt);
                StringBuilder sb = new StringBuilder();

                int num = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string title = dr["Title"].ToString();
                    string s3 = dr["Statement3"].ToString();

                    double totalValue = 0;
                    string f3 = ReplaceFormula(s3, ref totalValue);//1
                    
                    if (totalValue != 0)
                    {
                        isAllZero = false;

                        sb.Append(string.Format(title, num++));
                        sb.Append("|");
                        sb.Append(f3);
                        sb.Append("|");
                    }
                }

                if (isAllZero)
                {
                    sb.Append("无纳税调减项目。");
                }

                string sbString = sb.ToString();
                string ssbgMarkValue = sbString.EndsWith("|") ? sbString.Remove(sbString.Length - 1) : sbString;
                m_markTaxReport1.Value = ssbgMarkValue;
                m_ResultMarks.Add(m_markTaxReport1);
            }
            #endregion
        }

        string ReplaceFormula(string s, ref double totalValue)
        {
            int lastEndIndex = ParseFormula(0, ref s, ref totalValue);
            while (lastEndIndex != -1)
            {
                lastEndIndex = ParseFormula(lastEndIndex, ref s, ref totalValue);
            }

            return s;
        }

        int ParseFormula(int beginIndex, ref string s, ref double totalValue)
        {
            int startIndex = s.IndexOf('[', beginIndex);
            if (startIndex == -1)
                return -1;

            int endIndex = s.IndexOf(']', startIndex);
            if (endIndex == -1)
                return -1;
            else
            {
                string replaceHolder = s.Substring(startIndex, endIndex - startIndex + 1);
                string formula = replaceHolder.Substring(1, replaceHolder.Length - 2);
                double value = GetFromulaValue(formula);
                //totalValue += value;
                totalValue = value;
                string temp = " " + value.ToString("N") + " ";
                s = s.Replace(replaceHolder, temp);
                return endIndex + (temp.Length - replaceHolder.Length);
            }
        }

        double GetFromulaValue(string formula)
        {
            double value;
            string str = DAL.GetOtherMarkValue(formula, m_projectId);
            if (double.TryParse(str, out value))
                return value;
            else
                return 0d;
        }

        #endregion

        private string[] ParseTaxReportFormula(string formula, DataTable dt)
        {
            string[] array = new string[2];
            string formulaValue = formula;
            char[] operators = new char[]{'+', '-', '*', '/'};
            string[] genes = formula.Split(operators);
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < genes.Length; i++)
            {
                DataRow dr = dt.Rows.Find(new object[] { genes[i], i + 1 });
                if (dr != null)
                {
                    string markName = dr["Name"].ToString();
                    string markValue = DAL.GetOtherMarkValue(markName, m_projectId);
                    //string markTitleBegin = dr["TitleBegin"].ToString();
                    //string formula = dr["Formula"].ToString();
                    //string markTitleEnd = dr["TitleEnd"].ToString();
                    //int number = Convert.ToInt32(dr["Number"]);        //NOTE:暂时未用
                    //int order = Convert.ToInt32(dr["Order"]);
                    //bool bEnd = Convert.ToInt32(dr["BEnd"]) == 0 ? false : true;

                    if (!String.IsNullOrEmpty(markValue))
                    {
                        sb.Append(String.Format(@"|{0})、{1}:{2};", i + 1, markName, markValue));
                        formulaValue = formulaValue.Replace(genes[i], markValue);
                    }
                }//end if
            }

            array[0] = formulaValue;
            array[1] = sb.ToString();
            return array;
        }

        /// <summary>
        /// 不包含"NC,NM,QC,QM,JF,DF,JL,DL"中的任一值时返回false
        /// </summary>
        /// <param name="partB"></param>
        /// <returns></returns>
        private static bool CheckIsBalSubMark(string partB)
        {
            foreach (string tag in s_BalSubMarkTags)
            {
                if (partB.Contains(tag))
                    return true;
            }

            return false;
        }

        private static bool CheckIsAllZero(List<string> list)
        {
            bool isAllZero = true;
            foreach (string str in list)
            {
                if (double.Parse(str) != 0)
                {
                    isAllZero = false;
                    break;
                }
            }

            return isAllZero;
        }
    }
}
