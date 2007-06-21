using System;
using System.Collections.Generic;
using System.Text;

namespace AuditOfficeLibrary
{
    public class Mark
    {
        private string _formula;
        public string Formula
        {
            get { return _formula; }
            set { _formula = value; }
        }

        private int _x;
        public int X
        {
            get { return _x; }
        }

        private int _y;
        public int Y
        {
            get { return _y; }
        }

        private string _value = "";
        public string Value
        {
            set { _value = value; }
            get { return _value; }
        }

        private string _type = "";
        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }

        private string _markMean = "";

        public string MarkMean
        {
            get { return _markMean; }
            set { _markMean = value; }
        }

        private bool _available = true;

        public bool Available
        {
            get { return _available; }
            set { _available = value; }
        }

        private int _sheetIndex;

        public int SheetIndex
        {
            get { return _sheetIndex; }
            set { _sheetIndex = value; }
        }

        public Mark(string formula, int x, int y, int sheetIndex)
        {
            _formula = formula;
            _x = x;
            _y = y;
            _sheetIndex = sheetIndex;
        }
    }

    public class AnnotationTurple
    {
        private Dictionary<string, object> valueList = new Dictionary<string, object>();

        public AnnotationTurple()
        {
            foreach (string anoMarkPrefix in Ano.AnoCollection)
            {
                valueList.Add(anoMarkPrefix, null);
            }
        }

        public object this[string index]
        {
            get 
            {
                if (valueList.ContainsKey(index))
                {
                    return valueList[index];
                }
                return null;
            }
            set
            {
                if (valueList.ContainsKey(index))
                {
                    valueList[index] = value;
                }
            }
        }
    }


    public class Ano
    {
        private static Dictionary<string, Type> m_IndexDictionay = new Dictionary<string, Type>();
        private static Dictionary<string, string> m_IndexNameDictionay = new Dictionary<string, string>();
        private static List<string> m_AnoCollection = new List<string>();

        static Ano()
        {
            Type tString = typeof(string);
            Type tDouble = typeof(double);

            m_IndexDictionay.Add(NC, tDouble);
            m_IndexDictionay.Add(NM, tDouble);
            m_IndexDictionay.Add(JL, tDouble);
            m_IndexDictionay.Add(DL, tDouble);
            m_IndexDictionay.Add(VAL1, tDouble);
            m_IndexDictionay.Add(VAL2, tDouble);
            m_IndexDictionay.Add(TXT1, tString);
            m_IndexDictionay.Add(TXT2, tString);
            m_IndexDictionay.Add(TXT3, tString);
            m_IndexDictionay.Add(TXT4, tString);

            m_IndexNameDictionay.Add(NC, "年初数");
            m_IndexNameDictionay.Add(NM, "年末数");
            m_IndexNameDictionay.Add(JL, "借累");
            m_IndexNameDictionay.Add(DL, "贷累");
            m_IndexNameDictionay.Add(VAL1, "值一");
            m_IndexNameDictionay.Add(VAL2, "值二");
            m_IndexNameDictionay.Add(TXT1, "文本一");
            m_IndexNameDictionay.Add(TXT2, "文本二");
            m_IndexNameDictionay.Add(TXT3, "文本三");
            m_IndexNameDictionay.Add(TXT4, "文本四");

            m_AnoCollection.Add(AnoXM);
            m_AnoCollection.Add(AnoNC);
            m_AnoCollection.Add(AnoNM);
            m_AnoCollection.Add(AnoJL);
            m_AnoCollection.Add(AnoDL);
            m_AnoCollection.Add(AnoVAL1);
            m_AnoCollection.Add(AnoVAL2);
            m_AnoCollection.Add(AnoTXT1);
            m_AnoCollection.Add(AnoTXT2);
            m_AnoCollection.Add(AnoTXT3);
            m_AnoCollection.Add(AnoTXT4);
        }


        public static Dictionary<string, Type> IndexDictionay
        {
            get { return m_IndexDictionay; }
        } 

        public static List<string> AnoCollection
        {
            get { return m_AnoCollection; }
        }

        public static Dictionary<string, string> IndexNameDictionay
        {
            get { return Ano.m_IndexNameDictionay; }
            set { Ano.m_IndexNameDictionay = value; }
        }


        public static string TXT1
        {
            get { return "TXT1"; } 
        }

        public static string AnoTXT1
        {
            get { return AppendAnoPrefix(TXT1); }
        }

        public static string TXT2
        {
            get { return "TXT2"; }
        }

        public static string AnoTXT2
        {
            get { return AppendAnoPrefix(TXT2); }
        }

        public static string TXT3
        {
            get { return "TXT3"; }
        }

        public static string AnoTXT3
        {
            get { return AppendAnoPrefix(TXT3); }
        }

        public static string TXT4
        {
            get { return "TXT4"; }
        }

        public static string AnoTXT4
        {
            get { return AppendAnoPrefix(TXT4); }
        }

        public static string VAL1
        {
            get { return "VAL1"; }
        }

        public static string AnoVAL1
        {
            get { return AppendAnoPrefix(VAL1); }
        }

        public static string VAL2
        {
            get { return "VAL2"; }
        }

        public static string AnoVAL2
        {
            get { return AppendAnoPrefix(VAL2); }
        }

        public static string XM
        {
            get { return "XM"; }
        }

        public static string AnoXM
        {
            get { return AppendAnoPrefix(XM); }
        }

        public static string NC
        {
            get { return "NC"; }
        }

        public static string AnoNC
        {
            get { return AppendAnoPrefix(NC); }
        }

        public static string NM
        {
            get { return "NM"; }
        }

        public static string AnoNM
        {
            get { return AppendAnoPrefix(NM); }
        }

        public static string JL
        {
            get { return "JL"; }
        }

        public static string AnoJL
        {
            get { return AppendAnoPrefix(JL); }
        }

        public static string DL
        {
            get { return "DL"; }
        }

        public static string AnoDL
        {
            get { return AppendAnoPrefix(DL); }
        }

        public static string AppendAnoPrefix(string str)
        {
            return "ANO_" + str;
        }
    }
}
