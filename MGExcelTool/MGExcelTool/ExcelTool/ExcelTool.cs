using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Forms;

namespace MGExcelTool
{
    class ExcelTool
    {
        public static DataTable GetExcelToDataTableBySheet(string FileFullPath, string SheetName)
        {
            //string
            //strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 8.0; HDR=NO; IMEX=1'"; //此连接只能操作Excel2007之前(.xls)文件
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" +
           "data source=" + FileFullPath + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
            //此连接可以操作.xls与.xlsx文件
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter odda = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), conn);
            odda.Fill(ds, SheetName);
            conn.Close();
            return ds.Tables[0];
        }

        public static string[] GetExcelSheetNames(string FileFullPath)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;
            try
            {
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
                objConn = new OleDbConnection(strConn);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                string[] excelSheets = new string[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                return excelSheets;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public static DataTable[] GetExcelData(string FileFullPath)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;
            try
            {
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
                objConn = new OleDbConnection(strConn);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                DataTable[] dataTables = new DataTable[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    string sheetName = row["TABLE_NAME"].ToString();
                    DataSet ds = new DataSet();
                    OleDbDataAdapter odda = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", sheetName), objConn);
                    odda.Fill(ds, sheetName);
                    i++;
                }

                return dataTables;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public static Dictionary<string, DataTable> GetExcelDataDic(string FileFullPath)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;
            try
            {
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
                objConn = new OleDbConnection(strConn);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                Dictionary<string, DataTable> dataTables = new Dictionary<string, DataTable>();
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    string sheetName = row["TABLE_NAME"].ToString();
                    DataSet ds = new DataSet();
                    OleDbDataAdapter odda = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", sheetName), objConn);
                    odda.Fill(ds, sheetName);
                    if (sheetName.EndsWith("$"))
                    {
                        sheetName = sheetName.Remove(sheetName.Length - 1);
                    }
                    dataTables.Add(sheetName, ds.Tables[0]);
                    i++;
                }

                return dataTables;
            }
            catch (Exception e)
            {
                DebugHelper.Debug(e.ToString());
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }
    }
}
