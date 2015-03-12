using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MGExcelTool
{
    class CSExporter
    {
        //生成.cs文件
        public static void ExportToCSFile(object[] itemArray , string outCSFolderPath, string tableName)
        {
            try
            {
                //foreach (KeyValuePair<string, DataTable> kvp in dataTable)
                //{
                StringWriter sw = new StringWriter();

                //string tableName = kvp.Key;
                //DataTable dt = kvp.Value;
                //int rowCount = dt.Rows.Count;

                //Rows[0]注释，Rows[1]表字段，正式内容从2开始
                List<string> tableKeyList = new List<string>();
                //DataRow dr = dt.Rows[1];            //表key行
                //object[] itemArray = dr.ItemArray;
                sw.WriteLine("using System;");
                sw.WriteLine("");
                sw.WriteLine("class " + tableName + " : BaseConfigData");
                sw.WriteLine("{");
                for (int i = 0; i < itemArray.Length; i++)
                {
                    if (string.IsNullOrEmpty(itemArray[i].ToString()))
                    {
                        continue;
                    }

                    string[] strArrs = itemArray[i].ToString().Split(':');

                    string keyName = strArrs[0];
                    string keyType = strArrs[1];
                    if (keyType.Equals("id"))
                    {
                        sw.WriteLine("\t" + "public string " + keyName + ";");
                    }
                    else if (keyType.Equals("string"))
                    {
                        sw.WriteLine("\t" + "public string " + keyName + ";");
                    }
                    else if (keyType.Equals("int"))
                    {
                        sw.WriteLine("\t" + "public int " + keyName + ";");
                    }
                    else if (keyType.Equals("bool"))
                    {
                        sw.WriteLine("\t" + "public bool " + keyName + ";");
                    }
                    else if (keyType.Equals("float"))
                    {
                        sw.WriteLine("\t" + "public float " + keyName + ";");
                    }
                    else if (keyType.Equals("idArr"))
                    {
                        sw.WriteLine("\t" + "public string[] " + keyName + ";");
                    }
                    else if (keyType.Equals("refid"))
                    {
                        sw.WriteLine("\t" + "public string " + keyName + ";");
                    }
                    else
                    {
                        DebugHelper.DebugError("未知的数据类型，表名：" + tableName + " ，字段名： " + keyName + "，类型：" + keyType);
                    }
                }
                sw.WriteLine("}");

                string realInExcelFolderPath = Application.StartupPath + outCSFolderPath + "\\" + tableName + ".cs";
                File.WriteAllText(realInExcelFolderPath, sw.ToString(), Encoding.UTF8);
                sw.Close();

                DebugHelper.Debug("成功生成cs文件：" + realInExcelFolderPath);
                //}
            }
            catch (Exception e)
            {
                DebugHelper.Debug(e.ToString());
            }
        }
    }
}
