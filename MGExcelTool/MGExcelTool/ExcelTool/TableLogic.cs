using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MGExcelTool
{
    class TableLogic
    {
        static List<string> tableKeyList = new List<string>();

        //加入key/
        public static bool AddTableKey(string tableKey)
        {
            //符合主key逻辑/
            if (!tableKeyList.Contains(tableKey))
            {
                tableKeyList.Add(tableKey);
                return true;
            }

            return false;
        }

        //检查key是否符合外键逻辑/
        public static bool CheckTableKey(string tableKey)
        {
            return tableKeyList.Contains(tableKey);
        }
    }
}
