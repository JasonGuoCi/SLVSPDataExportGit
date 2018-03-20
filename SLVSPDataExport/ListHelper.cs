
using Microsoft.SharePoint;
using SLVSPDataExport.Utility;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLVSPDataExport
{
    public class ListHelper
    {

        public static bool IsListExist(string siteUrl, string listName)
        {
            bool isExist = false;
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    site.AllowUnsafeUpdates = true;
                    using (SPWeb web = site.OpenWeb())
                    {
                        web.AllowUnsafeUpdates = true;
                        SPListCollection lists = web.Lists;
                        foreach (SPList list in lists)
                        {
                            if (list.Title == listName)
                            {
                                isExist = true;
                            }
                        }
                        web.AllowUnsafeUpdates = false;
                    }
                    site.AllowUnsafeUpdates = false;
                }
            });
            return isExist;
        }

        public static DataTable GeChangedItems(DataRow[] rows, string siteUrl, string listName, string viewName, string logPath, string successLogPath)
        {
            DataTable dt = null;
            DataRow dataRow = null;
            Dictionary<string, string> listDictionary = new Dictionary<string, string>();
            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        bool isListExist = IsListExist(siteUrl, listName);
                        if (isListExist)
                        {
                            LogHelper.WriteLogSuccess("The list " + listName + " exist", successLogPath);
                            SPList list = web.Lists.TryGetList(listName);
                            if (list != null)
                            {
                                SPView view = list.Views[viewName];

                                int listItemCount = 0;
                                if (list.Items.Count > 0)
                                {
                                    SPListItemCollection items = list.GetItems(view);

                                    StringCollection viewFields = view.ViewFields.ToStringCollection();
                                    listItemCount = items.Count;
                                    dt = new DataTable();
                                    foreach (SPListItem item in items)
                                    {
                                        dataRow = dt.NewRow();
                                        for (int i = 0; i < rows.Length; i++)
                                        {
                                            //itemValue.add(item[rows[i][1]]);
                                            string fieldName = rows[i][1].ToString();
                                            dataRow[i] = item[fieldName];

                                        }
                                        dt.Rows.Add(dataRow);

                                    }
                                }
                                //foreach (SPListItem item in items)
                                //{
                                //    foreach (string fieldName in viewFields)
                                //    {
                                //        Console.WriteLine("{0} = {1}", fieldName, item[fieldName]);
                                //    }
                                //    Console.WriteLine();
                                //}
                            }
                        }
                        else
                        {
                            LogHelper.WriteLogSuccess("The list " + listName + " not exist", successLogPath);
                        }

                    }
                }
                #region withRunWithElevatedPrivileges
                /*
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    using (SPSite site = new SPSite(siteUrl))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            bool isListExist = IsListExist(siteUrl, listName);
                            if (isListExist)
                            {
                                LogHelper.WriteLogSuccess("The list " + listName + " exist", successLogPath);
                                SPList list = web.Lists.TryGetList(listName);
                                if (list != null)
                                {
                                    SPView view = list.Views[viewName];

                                    int listItemCount = 0;
                                    if (list.Items.Count > 0)
                                    {
                                        SPListItemCollection items = list.GetItems(view);

                                        StringCollection viewFields = view.ViewFields.ToStringCollection();
                                        listItemCount = items.Count;
                                        dt = new DataTable();
                                        foreach (SPListItem item in items)
                                        {
                                            dataRow = dt.NewRow();
                                            for (int i = 0; i < rows.Length; i++)
                                            {
                                                //itemValue.add(item[rows[i][1]]);
                                                string fieldName = rows[i][1].ToString();
                                                dataRow[i] = item[fieldName];

                                            }
                                            dt.Rows.Add(dataRow);

                                        }
                                    }
                                    //foreach (SPListItem item in items)
                                    //{
                                    //    foreach (string fieldName in viewFields)
                                    //    {
                                    //        Console.WriteLine("{0} = {1}", fieldName, item[fieldName]);
                                    //    }
                                    //    Console.WriteLine();
                                    //}
                                }
                            }
                            else
                            {
                                LogHelper.WriteLogSuccess("The list " + listName + " not exist", successLogPath);
                            }

                        }
                    }
                });
                 * */
                #endregion
                return dt;
            }
            catch (Exception ex)
            {

                LogHelper.WriteLog("Something wrong in GetChangedItems methed while retrieve the list items in the list " + listName + ", the error: " + ex.ToString(), logPath);
                return null;
            }

        }
    }
}
