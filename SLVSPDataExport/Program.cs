
using SLVSPDataExport.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLVSPDataExport
{
    class Program
    {
        private static string logPath;
        private static string logPathSuccess;
        private static string sourceSPListFields;
        private static string resultPath;
        private static string headerName = string.Empty;
        private static string line = string.Empty;
        private static string[] listArray = { "Project Basic List", "ApplicationArea", "Bulb Data", "Characteristics", "Component Data", "Delivery includes", "Device additional signs", "Device Data", "Dimension", "EEK Bulb", "Fixture Data", "Lamp type LED", "Lamp Type Socket", "LEDModules", "LEDStrips", "Location Area", "Light Distribution Curve", "Profile", "Quality Aspects", "Radio", "Recommended Bulb", "Accessories" };
        //private static int fieldsCount = 459;
        // private static int[] listFieldsCount = { 14, 11, 38, 20, 10, 18, 10, 49, 17, 2, 52, 49, 47, 29, 33, 19, 5, 3, 6, 6, 4, 14 };
        private static List<int> listFieldsCount = new List<int>();
        private static string siteUrl;
        private static string viewName;
        private static string userName;
        private static string pwd;
        private static string domain;
        static void Main(string[] args)
        {
            Console.WriteLine("Export Data Start...");
            logPath = ConfigurationManager.AppSettings["logPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
            logPathSuccess = ConfigurationManager.AppSettings["logPathSuccess"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
            sourceSPListFields = ConfigurationManager.AppSettings["sourceSPListFields"];
            resultPath = ConfigurationManager.AppSettings["resultPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".csv";
            siteUrl = ConfigurationManager.AppSettings["siteUrl"];
            viewName = ConfigurationManager.AppSettings["viewName"];
            userName = ConfigurationManager.AppSettings["userName"];
            pwd = ConfigurationManager.AppSettings["pwd"];
            domain = ConfigurationManager.AppSettings["domain"];
            Write2CSV();
            Console.WriteLine("Add finish..., please click any key to exist.");
            Console.ReadKey();
        }

        public static void Write2CSV()
        {
            DataTable sourceSPListFieldsDT = ExcelHelper.ExcelToDataTable(sourceSPListFields, true, logPath, logPathSuccess);
            for (int m = 0; m < listArray.Count(); m++)
            {
                string lstName = listArray[m];
                //qu de xiang tong list de ziduan
                DataRow[] rowsArray = sourceSPListFieldsDT.Select("SPList='" + lstName + "'");
                listFieldsCount.Add(rowsArray.Length);
            }
            if (sourceSPListFieldsDT != null && sourceSPListFieldsDT.Rows.Count > 0)
            {
                for (int i = 0; i < sourceSPListFieldsDT.Rows.Count; i++)
                {
                    //headerName += sourceSPListFieldsDT.Rows[i]["Header"].ToString() + ",";
                    headerName += sourceSPListFieldsDT.Rows[i]["Header"].ToString() + ";";
                    //for (int j = 0; j < sourceSPListFieldsDT.Columns.Count; j++)
                    //{
                    //    headerName += sourceSPListFieldsDT.Rows[i]["Header"].ToString() + ",";
                    //}
                }
                //headerName = headerName.TrimEnd(',');
                headerName = headerName.TrimEnd(';');
                WriteResult(headerName);
                LogHelper.WriteLogSuccess("Write header columns success", logPathSuccess);
                try
                {
                    //bianli listarray zhong de list
                    for (int m = 0; m < listArray.Count(); m++)
                    {
                        string lstName = listArray[m];
                        //string lstName = "Recommended Bulb";//for test
                        //qu de xiang tong list de ziduan
                        DataRow[] rowsArray = sourceSPListFieldsDT.Select("SPList='" + lstName + "'");
                        if (rowsArray.Length > 0)
                        {

                            DataTable dt = ListHelperCSOM.GeChangedItems(rowsArray, siteUrl, userName, pwd, domain, lstName, viewName, logPath, logPathSuccess);
                            //DataTable dt = ListHelper.GeChangedItems(rowsArray, siteUrl, listArray[m], viewName, logPath, logPathSuccess);


                            if (dt != null && dt.Rows.Count > 0)
                            {
                                int rowCount = dt.Rows.Count;//行数  
                                int columnCount = dt.Columns.Count;//列数  
                                if (m == 0)
                                {

                                    for (int i = 0; i < rowCount; i++)
                                    {
                                        for (int j = 0; j < columnCount; j++)
                                        {
                                            //line = line + dt.Rows[i][j].ToString() + ",";
                                            line = line + dt.Rows[i][j].ToString() + ";";
                                        }
                                        //line = line.TrimEnd(',');
                                        line = line.TrimEnd(';');
                                        WriteResult(line);
                                        line = string.Empty;
                                    }
                                }
                                else
                                {
                                    string newline = line + dt.Rows[0][0].ToString() + ";;" + dt.Rows[0][1];
                                    int semicolonCount = 0;
                                    for (int i = 0; i < m; i++)
                                    {
                                        semicolonCount = semicolonCount + listFieldsCount[i];
                                    }
                                    string simicolonSeparator = string.Empty;
                                    for (int j = 0; j < semicolonCount - 2; j++)
                                    {
                                        //line = line + ",";
                                        simicolonSeparator = simicolonSeparator + ";";
                                    }
                                    //simicolonSeparator = line + simicolonSeparator;

                                    for (int i = 0; i < rowCount; i++)
                                    {
                                        line = newline + simicolonSeparator;
                                        for (int j = 2; j < columnCount; j++)
                                        {
                                            //line = line + dt.Rows[i][j].ToString() + ",";

                                            line = line + dt.Rows[i][j].ToString() + ";";
                                        }
                                        //line = line.TrimEnd(',');
                                        line = line.TrimEnd(';');
                                        WriteResult(line);
                                        line = string.Empty;
                                    }
                                }

                            }



                        }
                    }
                    LogHelper.WriteLogSuccess("Write data to csv success", logPathSuccess);
                }
                catch (Exception ex)
                {

                    LogHelper.WriteLog("Something wrong in the method Write2CSV: " + ex.ToString(), logPath);
                }

                //for (int i = 1; i < sourceSPListFieldsDT.Rows.Count; i++)
                //{

                //foreach (var list in listArray)
                //{
                //    DataRow[] rows = sourceSPListFieldsDT.Select("SPList= '" + list + "'");
                //    if (rows.Length > 0)
                //    {
                //        DataTable dt = GetSPListItems(rows, list);
                //    }
                //}
                //}
            }
        }

        public static void WriteResult(string text)
        {
            FileStream fs = new FileStream(resultPath, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.WriteLine(text);
            sw.Close();
            fs.Close();
        }
    }
}
