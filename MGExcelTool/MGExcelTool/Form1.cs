using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MGExcelTool;

namespace MGExcelTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DebugHelper.debugShowBox = DebugTextBox;
        }

        private string inExcelFolderPath = "\\ConfigExcels";
        private string outXmlFolderPath = "\\Output\\ConfigXmls";
        private string outCSFolderPath = "\\Output\\ConfigCS";

        private void button1_Click(object sender, EventArgs e)
        {
            DebugHelper.Clear();

            string realInExcelFolderPath = Application.StartupPath + inExcelFolderPath;

            if (Directory.Exists(realInExcelFolderPath))
            {
                DebugHelper.Debug("Find directory--->" + realInExcelFolderPath);

                string[] files = Directory.GetFiles(realInExcelFolderPath);
                for (int i = 0; i < files.Length;i++)
                {
                    string fileName = files[i];
                    if (fileName.EndsWith(".xlsx") || fileName.EndsWith(".xls"))
                    {
                        Dictionary<string, DataTable> dataTable = ExcelTool.GetExcelDataDic(fileName);
                        XmlExporter.ExportToXml(dataTable, fileName, outXmlFolderPath ,outCSFolderPath);
                    }
                    else
                    {
                        DebugHelper.Debug(fileName + "--> is not excel file!");
                    }
                }
            }
            else
            {
                DebugHelper.Debug("Directory not exits--->" + realInExcelFolderPath);
            }
        }
    }
}
