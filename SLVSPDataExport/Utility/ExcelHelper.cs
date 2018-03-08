using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace SLVSPDataExport.Utility
{
    public class ExcelHelper
    {
        /// <summary>  
        /// 将excel导入到datatable  
        /// </summary>  
        /// <param name="filePath">excel路径</param>  
        /// <param name="isColumnName">第一行是否是列名</param>  
        /// <returns>返回datatable</returns>  
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName, string logPath, string successLogPath)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet  
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行  
                                int cellCount = firstRow.LastCellNum;//列数  

                                //构建datatable的列  
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行  
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                            //switch (cell.CellType)
                                            //{
                                            //    case CellType.Blank:
                                            //        dataRow[j] = "";
                                            //        break;
                                            //    case CellType.Numeric:
                                            //        dataRow[j] = cell.DateCellValue;
                                            //        //short format = cell.CellStyle.DataFormat;
                                            //        ////对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                            //        //if (format == 14 || format == 31 || format == 57 || format == 58)
                                            //        //    dataRow[j] = cell.DateCellValue;
                                            //        //else
                                            //        //    dataRow[j] = cell.NumericCellValue;
                                            //        break;
                                            //    case CellType.String:
                                            //        dataRow[j] = cell.StringCellValue;
                                            //        break;
                                            //}
                                            switch (cell.CellType)
                                            {
                                                case CellType.String:
                                                    string str = row.GetCell(j).StringCellValue;
                                                    if (str != null && str.Length > 0)
                                                    {
                                                        dataRow[j] = str.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                                    {
                                                        dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                                    }
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                case CellType.Formula:
                                                    switch (row.GetCell(j).CachedFormulaResultType)
                                                    {
                                                        case CellType.String:
                                                            string strFORMULA = row.GetCell(j).StringCellValue;
                                                            if (strFORMULA != null && strFORMULA.Length > 0)
                                                            {
                                                                dataRow[j] = strFORMULA.ToString();
                                                            }
                                                            else
                                                            {
                                                                dataRow[j] = null;
                                                            }
                                                            break;
                                                        case CellType.Numeric:
                                                            dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                            break;
                                                        case CellType.Boolean:
                                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                            break;
                                                        case CellType.Error:
                                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                            break;
                                                        default:
                                                            dataRow[j] = "";
                                                            break;
                                                    }
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;

                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                LogHelper.WriteLog("ExcelToDataTable method error:" + ex.ToString(), logPath);
                return null;
            }
        }

        /// <summary>
        /// wirte dt to excel
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool DataTableToExcel(DataTable dt, string exportPath)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new XSSFWorkbook();//用XSSF==>xlsx后缀，而HSSF==>xls后缀
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  

                    //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    if (File.Exists(exportPath))
                    {
                        File.Delete(exportPath);
                    }

                    using (fs = File.OpenWrite(exportPath))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }


                    //FileStream xlsfile = new FileStream(exportPath, FileMode.Create);
                    //workbook.Write(xlsfile);
                    //xlsfile.Close();
                    //result = true;
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }
        }


        /// <summary>  
        /// 将excel导入到hashtable
        /// </summary>  
        /// <param name="filePath">excel路径</param>  
        /// <param name="isColumnName">第一行是否是列名</param>  
        /// <returns>返回datatable</returns>  
        public static Hashtable ExcelToHashTable(string filePath, bool isColumnName, string logPath, string successLogPath)
        {
            Hashtable hashTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet  
                        hashTable = new Hashtable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行  
                                int cellCount = firstRow.LastCellNum;//列数  

                                //构建datatable的列  
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取  

                                }


                                //填充行  
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null)
                                    {
                                        LogHelper.WriteLog("The row " + i + " is null", logPath);
                                        continue;
                                    }
                                    if (row.GetCell(0) == null)
                                    {
                                        LogHelper.WriteLog("The cell Code of row " + i + " is null/empty", logPath);
                                        continue;
                                    }
                                    if (hashTable.Contains(row.GetCell(0).StringCellValue))
                                    {
                                        LogHelper.WriteLog("The key " + row.GetCell(0).StringCellValue + " has existed", logPath);
                                        continue;
                                    }
                                    string val = string.Empty;
                                    if (row.GetCell(1) == null)
                                    {
                                        val = "";
                                    }
                                    else
                                    {
                                        val = row.GetCell(1).StringCellValue;
                                    }
                                    hashTable.Add(row.GetCell(0).StringCellValue, val);
                                }
                            }
                        }
                    }
                }
                return hashTable;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                LogHelper.WriteLog("ExcelToHashtable method error:" + ex.ToString(), logPath);
                return null;
            }
        }

    }
}
