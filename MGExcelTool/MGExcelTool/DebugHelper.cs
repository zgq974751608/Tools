
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing;

namespace MGExcelTool
{
    class DebugHelper
    {
        public static RichTextBox debugShowBox;

        private static string logFilePath = Application.StartupPath + "\\log.txt";

        private static StringWriter sw = new StringWriter();

        public static void Debug(Object obj)
        {
            if (debugShowBox != null)
            {
                sw.WriteLine(obj);
                string content = obj.ToString() + ";\r\n";
                debugShowBox.SelectionColor = Color.Black;
                debugShowBox.SelectedText = content;
                //debugShowBox.AppendText(obj.ToString() + "\r\n");
                //debugShowBox.Text = sw.ToString();
            }
            else
            {
                sw.WriteLine("Debug Textbox not inited!");
            }
            WriteLog();
        }

        public static void DebugError(Object obj)
        {
            if (debugShowBox != null)
            {
                sw.WriteLine(obj);
                string content = obj.ToString() + ";\r\n";
                debugShowBox.SelectionColor = Color.Red;
                debugShowBox.SelectedText = content;
                //debugShowBox.AppendText(obj.ToString() + "\r\n");
                //debugShowBox.Text = sw.ToString();
            }
            else
            {
                sw.WriteLine("Debug Textbox not inited!");
            }
            WriteLog();
        }

        public static void Clear()
        {
            if (sw != null)
            {
                sw = null;
                sw = new StringWriter();
                debugShowBox.Text = "";
            }
        }

        static void WriteLog()
        {
            File.WriteAllText(logFilePath, sw.ToString());
        }
    }
}
