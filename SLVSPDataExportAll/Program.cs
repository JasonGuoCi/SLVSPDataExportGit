using SLVSPDataExportAll.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLVSPDataExportAll
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
            //logPath = ConfigurationManager.AppSettings["logPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".txt";
            logPath = ConfigurationManager.AppSettings["logPath"] + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss") + ".txt";
            logPathSuccess = ConfigurationManager.AppSettings["logPathSuccess"] + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss") + ".txt";
            sourceSPListFields = ConfigurationManager.AppSettings["sourceSPListFields"];
            resultPath = ConfigurationManager.AppSettings["resultPath"] + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss") + ".csv";
            siteUrl = ConfigurationManager.AppSettings["siteUrl"];
            viewName = ConfigurationManager.AppSettings["viewName"];
            userName = ConfigurationManager.AppSettings["userName"];
            pwd = ConfigurationManager.AppSettings["pwd"];
            domain = ConfigurationManager.AppSettings["domain"];
            LogHelper.WriteLogSuccess("Export begin " + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss"), logPathSuccess);
            /* for 

            ListHelperCSOM.GetItemById(siteUrl, userName, pwd, domain);
             * 
             * test*/
            //ListHelperCSOM.GetItemById(siteUrl, userName, pwd, domain);
            Write2CSV();
            LogHelper.WriteLogSuccess("Export end " + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss"), logPathSuccess);
            Console.WriteLine("Add finish..., please click any key to exist.");
            Console.ReadKey();
        }

        public static void Write2CSV()
        {
            DataTable sourceSPListFieldsDT = ExcelHelper.ExcelToDataTable(sourceSPListFields, true, logPath, logPathSuccess);
            DataTable dtAll = new DataTable();
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
                    if (!dtAll.Columns.Contains(sourceSPListFieldsDT.Rows[i][1].ToString()))
                    {
                        dtAll.Columns.Add(sourceSPListFieldsDT.Rows[i][1].ToString(), Type.GetType("System.String"));
                    }
                }
                //headerName = headerName.TrimEnd(',');
                headerName = headerName.TrimEnd(';');
                WriteResult(headerName);
                LogHelper.WriteLogSuccess("Write header columns success", logPathSuccess);

                try
                {
                    for (int m = 0; m < listArray.Count(); m++)
                    {
                        string lstName = listArray[m];
                        //qu de xiang tong list zhong de ziduan
                        DataRow[] fields = sourceSPListFieldsDT.Select("SPList='" + lstName + "'");
                        if (fields.Length > 0)
                        {
                            DataTable dt = ListHelperCSOM.GetItems(fields, siteUrl, userName, pwd, domain, lstName, viewName, logPath, logPathSuccess);

                            #region export each sku in one line
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                int rowCount = dt.Rows.Count;
                                int columnCount = dt.Columns.Count;
                                for (int i = 0; i < rowCount; i++)
                                {
                                    DataRow[] dtAllRowsArray = dtAll.Select("[Material No]='" + dt.Rows[i]["Material No"] + "'");
                                    //if material no exist
                                    if (dtAllRowsArray.Length > 0)
                                    {
                                        foreach (DataRow dtAllRowsArrayRow in dtAllRowsArray)
                                        {
                                            for (int j = 0; j < fields.Length; j++)
                                            {
                                                dtAllRowsArrayRow[fields[j][1].ToString()] = dt.Rows[i][fields[j][1].ToString()];
                                            }

                                        }

                                    }
                                    else//material no not exists
                                    {
                                        if (dt.Rows[i][0].ToString() == "")
                                        {
                                            continue;
                                        }
                                        DataRow dtAllNewRow = dtAll.NewRow();
                                        dtAllNewRow["Material No"] = dt.Rows[i]["Material No"].ToString();
                                        dtAllNewRow["Project Article No"] = dt.Rows[i]["Project Article No"].ToString();
                                        for (int j = 0; j < fields.Length; j++)
                                        {
                                            dtAllNewRow[fields[j][1].ToString()] = dt.Rows[i][fields[j][1].ToString()];
                                        }
                                        dtAll.Rows.Add(dtAllNewRow);

                                    }

                                }

                            }

                            #endregion

                            #region get all changed data to csv
                            /*
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                int rowCount = dt.Rows.Count;//行数  
                                int columnCount = dt.Columns.Count;//列数  
                                if (m == 0)
                                {

                                    for (int i = 0; i < rowCount; i++)
                                    {
                                        if (dt.Rows[0][0].ToString() == "")
                                        {
                                            continue;
                                        }
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


                                    for (int i = 0; i < rowCount; i++)
                                    {
                                        if (dt.Rows[i][0].ToString() == "")
                                        {
                                            continue;
                                        }
                                        string newline = line + dt.Rows[i][0].ToString() + ";;" + dt.Rows[i][1];
                                        int semicolonCount = 0;
                                        for (int y = 0; y < m; y++)
                                        {
                                            semicolonCount = semicolonCount + listFieldsCount[y];
                                        }
                                        string simicolonSeparator = string.Empty;
                                        for (int j = 0; j < semicolonCount - 2; j++)
                                        {
                                            //line = line + ",";
                                            simicolonSeparator = simicolonSeparator + ";";
                                        }
                                        //simicolonSeparator = line + simicolonSeparator;

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

                            }*/
                            #endregion
                        }
                    }
                    string csvRow = string.Empty;
                    for (int j = 0; j < dtAll.Rows.Count; j++)
                    {
                        if (dtAll.Rows[j]["Material No"].ToString() == "")
                        {
                            continue;
                        }
                        for (int k = 0; k < sourceSPListFieldsDT.Rows.Count; k++)
                        {
                            if (sourceSPListFieldsDT.Rows[k][1].ToString() == "")
                            {
                                continue;
                            }
                            string csvFieldValue = dtAll.Rows[j][sourceSPListFieldsDT.Rows[k][1].ToString()] == null ? "" : dtAll.Rows[j][sourceSPListFieldsDT.Rows[k][1].ToString()].ToString();
                            csvRow += csvFieldValue + ";";
                        }
                        csvRow = csvRow.TrimEnd(';');
                        WriteResult(csvRow);
                        csvRow = string.Empty;
                    }
                    LogHelper.WriteLogSuccess("Write data to csv success", logPathSuccess);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog("Something wrong in the method Write2CSV: " + ex.ToString(), logPath);
                }
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
