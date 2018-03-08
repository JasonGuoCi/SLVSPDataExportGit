using Microsoft.SharePoint.Client;
using SLVSPDataExport.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SLVSPDataExport
{
    public class ListHelperCSOM
    {
        public static bool IsListExist(string siteUrl, string listName, string userName, string pwd, string domain)
        {
            ClientContext clientContext = new ClientContext(siteUrl);
            clientContext.Credentials = new NetworkCredential(userName, pwd, domain);

            ListCollection listCollection = clientContext.Web.Lists;
            clientContext.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listName));
            clientContext.ExecuteQuery();

            if (listCollection.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }


        public static DataTable GeChangedItems(DataRow[] rows, string siteUrl, string userName, string pwd, string domain, string listName, string viewName, string logPath, string successLogPath)
        {
            DataTable dt = new DataTable();
            if (listName != "Project Basic List")
            {
                dt.Columns.Add("Material No", Type.GetType("System.String"));
                dt.Columns.Add("Project Article No", Type.GetType("System.String"));
            }

            for (int i = 0; i < rows.Length; i++)
            {
                if (!dt.Columns.Contains(rows[i][1].ToString()))
                {
                    dt.Columns.Add(rows[i][1].ToString(), Type.GetType("System.String"));
                }
                //dt.Columns.Add(rows[i][1].ToString(), Type.GetType("System.String"));
            }
            DataRow dataRow = null;
            Dictionary<string, string> listDictionary = new Dictionary<string, string>();
            int total = 0;
            try
            {
                bool isListExist = IsListExist(siteUrl, listName, userName, pwd, domain);
                if (isListExist)
                {
                    LogHelper.WriteLogSuccess("The list " + listName + " exist", successLogPath);
                    ClientContext clientContext = new ClientContext(siteUrl);
                    clientContext.Credentials = new NetworkCredential(userName, pwd, domain);

                    ListCollection lists = clientContext.Web.Lists;
                    IEnumerable<List> results = clientContext.LoadQuery<List>(lists.Where(lst => lst.Title == listName));
                    clientContext.ExecuteQuery();
                    List list = results.FirstOrDefault();
                    if (list == null)
                    {
                        LogHelper.WriteLog("A list named " + listName + " does not exist. Press any key to exit...", logPath);
                        return null;
                    }

                    //List list = clientContext.Web.Lists.GetByTitle(listName);
                    //clientContext.ExecuteQuery();
                    ChangeQuery query = new ChangeQuery(false, false);
                    //query.FetchLimit = true;
                    query.Item = true;
                    query.Add = true;
                    query.Update = true;

                    query.ChangeTokenStart = new ChangeToken();
                    query.ChangeTokenStart.StringValue = string.Format("1;3;{0};{1};-1", list.Id.ToString(), DateTime.Now.AddDays(-10).ToUniversalTime().Ticks.ToString());

                    ChangeCollection changes = list.GetChanges(query);
                    clientContext.Load(changes);
                    clientContext.ExecuteQuery();

                    total += changes.Count;
                    List<int> itemIds = new List<int>();
                    foreach (ChangeItem changeItem in changes)
                    {
                        if (!itemIds.Contains(changeItem.ItemId))
                        {
                            itemIds.Add(changeItem.ItemId);
                        }

                    }
                    foreach (int itemId in itemIds)
                    {
                        //int itemId = changeItem.ItemId;
                        ListItem currentItem = list.GetItemById(itemId);
                        clientContext.Load(currentItem);
                        clientContext.ExecuteQuery();

                        dataRow = dt.NewRow();
                        if (currentItem != null)
                        {

                            if (listName == "Recommended Bulb" || listName == "Accessories")
                            {
                                Field skuField = list.Fields.GetByTitle("Material No");
                                Field articleNo = list.Fields.GetByTitle("Project Article No");
                                clientContext.Load(skuField);
                                clientContext.Load(articleNo);
                                clientContext.ExecuteQuery();
                                string skuFieldInternalName = skuField.InternalName;
                                string articleNoInternalName = articleNo.InternalName;

                                string skuFieldValue = string.Empty;
                                if (skuField.TypeDisplayName.Contains("lookup") || skuField.TypeDisplayName.Contains("Nachschlagen"))
                                {
                                    FieldLookupValue childIdField = currentItem[skuFieldInternalName] as FieldLookupValue;
                                    //string lookupValues = string.Empty;

                                    if (childIdField != null)
                                    {

                                        skuFieldValue = childIdField.LookupValue;
                                        //var childId_Id = lookupValue.LookupId;

                                    }
                                    //dataRow[fieldDisName] = lookupValues;
                                }
                                else
                                {
                                    skuFieldValue = currentItem[skuFieldInternalName] == null ? "" : currentItem[skuFieldInternalName].ToString();
                                }
                                dataRow["Material No"] = skuFieldValue;

                                string articleNoFieldValue = string.Empty;
                                if (articleNo.TypeDisplayName.Contains("lookup") || articleNo.TypeDisplayName.Contains("Nachschlagen"))
                                {
                                    FieldLookupValue childIdField = currentItem[articleNoInternalName] as FieldLookupValue;
                                    //string lookupValues = string.Empty;

                                    if (childIdField != null)
                                    {

                                        articleNoFieldValue = childIdField.LookupValue;
                                        //var childId_Id = lookupValue.LookupId;

                                    }
                                    //dataRow[fieldDisName] = lookupValues;
                                }
                                else
                                {
                                    articleNoFieldValue = currentItem[articleNoInternalName] == null ? "" : currentItem[articleNoInternalName].ToString();
                                }

                                dataRow["Project Article No"] = articleNoFieldValue;

                                string[] fields = rows[0][1].ToString().Split('&');
                                dataRow[rows[0][1].ToString()] = string.Empty;
                                for (int i = 0; i < fields.Length; i++)
                                {
                                    Field field = list.Fields.GetByTitle(fields[i].TrimStart().TrimEnd());
                                    clientContext.Load(field);
                                    clientContext.ExecuteQuery();
                                    string fieldInternalName = field.InternalName;

                                    dataRow[rows[0][1].ToString()] = dataRow[rows[0][1].ToString()].ToString() + "," + currentItem[fieldInternalName];
                                }
                                dataRow[rows[0][1].ToString()] = dataRow[rows[0][1].ToString()].ToString().TrimStart(',').TrimEnd(',');
                            }
                            else
                            {
                                for (int i = 0; i < rows.Length; i++)
                                {

                                    if (listName != "Project Basic List")
                                    {
                                        Field skuField = list.Fields.GetByTitle("Material No");
                                        Field articleNo = list.Fields.GetByTitle("Project Article No");
                                        clientContext.Load(skuField);
                                        clientContext.Load(articleNo);
                                        clientContext.ExecuteQuery();
                                        string skuFieldInternalName = skuField.InternalName;
                                        string articleNoInternalName = articleNo.InternalName;

                                        string skuFieldValue = string.Empty;
                                        if (skuField.TypeDisplayName.Contains("lookup") || skuField.TypeDisplayName.Contains("Nachschlagen"))
                                        {
                                            FieldLookupValue childIdField = currentItem[skuFieldInternalName] as FieldLookupValue;
                                            //string lookupValues = string.Empty;

                                            if (childIdField != null)
                                            {

                                                skuFieldValue = childIdField.LookupValue;
                                                //var childId_Id = lookupValue.LookupId;

                                            }
                                            //dataRow[fieldDisName] = lookupValues;
                                        }
                                        else
                                        {
                                            skuFieldValue = currentItem[skuFieldInternalName] == null ? "" : currentItem[skuFieldInternalName].ToString();
                                        }
                                        dataRow["Material No"] = skuFieldValue;

                                        string articleNoFieldValue = string.Empty;
                                        if (articleNo.TypeDisplayName.Contains("lookup") || articleNo.TypeDisplayName.Contains("Nachschlagen"))
                                        {
                                            FieldLookupValue childIdField = currentItem[articleNoInternalName] as FieldLookupValue;
                                            //string lookupValues = string.Empty;

                                            if (childIdField != null)
                                            {

                                                articleNoFieldValue = childIdField.LookupValue;
                                                //var childId_Id = lookupValue.LookupId;

                                            }
                                            //dataRow[fieldDisName] = lookupValues;
                                        }
                                        else
                                        {
                                            articleNoFieldValue = currentItem[articleNoInternalName] == null ? "" : currentItem[articleNoInternalName].ToString();
                                        }

                                        dataRow["Project Article No"] = articleNoFieldValue;
                                    }

                                    string fieldDisName = rows[i][1].ToString();
                                    Field field = list.Fields.GetByTitle(fieldDisName);
                                    clientContext.Load(field);
                                    clientContext.ExecuteQuery();
                                    string fieldInternalName = field.InternalName;

                                    bool flag = field.TypeDisplayName.Contains("lookup") || field.TypeDisplayName.Contains("Nachschlagen");
                                    if (flag)
                                    {
                                        FieldLookupValue childIdField = currentItem[fieldInternalName] as FieldLookupValue;
                                        string lookupValues = string.Empty;

                                        if (childIdField != null)
                                        {

                                            lookupValues = childIdField.LookupValue;
                                            //var childId_Id = lookupValue.LookupId;

                                        }
                                        dataRow[fieldDisName] = lookupValues;
                                    }
                                    else
                                    {
                                        dataRow[fieldDisName] = currentItem[fieldInternalName] == null ? "" : currentItem[fieldInternalName].ToString();
                                    }

                                }
                            }

                        }

                        dt.Rows.Add(dataRow);
                    }

                }

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
