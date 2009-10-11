using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace AuditOfficeLibrary
{
    public class Log
    {
        public static void Write(string message)
        {
            Debug.WriteLine(message); 
            //using (StreamWriter sw = File.AppendText(@"D:\Log.txt"))
            //{
            //    sw.WriteLine(message);
            //}
        }

        public static void Write(string message, bool showDialog)
        {
            Write(message);
            if (showDialog)
            {
                MessageBox.Show(message);
            }
        }
    }
}
