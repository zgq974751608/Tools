using System.Collections.Generic;
using System.Data;
using System.Xml;
using System.Windows.Forms;
using System;


namespace MGExcelTool
{
    class XmlExporter
    {
        //导出为Xml
        public static void ExportToXml(Dictionary<string, DataTable> dataTable, string FileFullPath, string outXmlPath, string outCSFolderPath)
        {
            try
            {
                foreach (KeyValuePair<string, DataTable> kvp in dataTable)
                {
                    string tableName = kvp.Key;
                    DataTable dt = kvp.Value;
                    int rowCount = dt.Rows.Count;

                    //Rows[0]注释，Rows[1]表字段，正式内容从2开始
                    List<string> tableKeyList = new List<string>();
                    List<string> tableKeyTypeList = new List<string>();
                    DataRow dr = dt.Rows[1];            //表key行
                    object[] itemArray = dr.ItemArray;
                    CSExporter.ExportToCSFile(itemArray, outCSFolderPath, tableName);
                    for (int i = 0; i < itemArray.Length; i++)
                    {
                        string[] strArrs = itemArray[i].ToString().Split(':');
                        string keyName = strArrs[0];
                        string keyType = strArrs.Length > 1 ? strArrs[1] : "";

                        if (tableKeyTypeList.Contains(keyType)){
                            DebugHelper.DebugError("表逻辑错误，只能有一个id");
                            return;
                        }
                        tableKeyTypeList.Add(keyType);
                        tableKeyList.Add(keyName);
                    }

                    XmlDocument doc = new XmlDocument();
                    XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "utf-8", null);
                    doc.AppendChild(dec);
                    XmlElement root = doc.CreateElement("Table");

                    for (int j = 2; j < rowCount; j++)
                    {
                        DataRow tempDataRow = dt.Rows[j];
                        object[] tempItemArray = tempDataRow.ItemArray;

                        XmlElement element = doc.CreateElement("TableRow");
                        for (int k = 0; k < tempItemArray.Length; k++)
                        {
                            string context = tempItemArray[k].ToString();
                            string keyName = tableKeyList[k];
                            if (string.IsNullOrEmpty(keyName))
                            {
                                //注释列，忽略/
                                continue;
                            }

                            string keyType = tableKeyTypeList[k];
                            //id或者refid的时候，检查key值逻辑/
                            if (keyType.Equals("id"))
                            {
                                bool correctkey = TableLogic.AddTableKey(context);
                                if (!correctkey)
                                {
                                    DebugHelper.DebugError("key值重复，请修改后重新运行，文件名：" + FileFullPath + "，表名：" + tableName + " ，字段名： " + keyName + " ,行号 : " + j);
                                    return;
                                }
                            }
                            else if (keyType.Equals("refid"))
                            {
                                bool correctForeignkey = TableLogic.CheckTableKey(context);
                                if (!correctForeignkey)
                                {
                                    DebugHelper.DebugError("key值不存在，请修改后重新运行，文件名：" + FileFullPath + "，表名：" + tableName + " ，字段名： " + keyName + " ,行号 : " + j);
                                    return;
                                }
                            }
                            else if (keyType.Equals("idArr"))
                            {
                                string[] contexts = context.Split(';');
                                for (int n = 0; n < contexts.Length;n++)
                                {
                                    bool correctForeignkey = TableLogic.CheckTableKey(contexts[n]);
                                    if (!correctForeignkey)
                                    {
                                        DebugHelper.DebugError("key值不存在，请修改后重新运行，文件名：" + FileFullPath + "，表名：" + tableName + " ，字段名： " + keyName + " ,行号 : " + j);
                                        return;
                                    }
                                }
                            }
                            element.SetAttribute(tableKeyList[k], context);
                        }
                        root.AppendChild(element);
                    }

                    doc.AppendChild(root);
                    string realInExcelFolderPath = Application.StartupPath + outXmlPath + "\\" + tableName + ".xml";
                    doc.Save(@realInExcelFolderPath);

                    DebugHelper.Debug("成功生成xml文件：" + realInExcelFolderPath);
                }
            } catch (Exception e){
                DebugHelper.Debug(e.ToString());
            }
        }
    }
}
