using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SLVSPDataExport.Utility
{
    public class LogHelper
    {
        /// <summary>
        /// write exception/error log
        /// </summary>
        /// <param name="text"></param>
        /// <param name="logPath"></param>
        public static void WriteLog(string text, string logPath)
        {
            FileStream fs = new FileStream(logPath, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.Default);
            sw.WriteLine(text);
            sw.Close();
            fs.Close();
        }
        /// <summary>
        /// write success log
        /// </summary>
        /// <param name="text"></param>
        /// <param name="logPathSuccess"></param>
        public static void WriteLogSuccess(string text, string logPathSuccess)
        {
            FileStream fs = new FileStream(logPathSuccess, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.Default);
            sw.WriteLine(text);
            sw.Close();
            fs.Close();
        }
    }
}
