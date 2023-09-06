using Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic
{
    public class ExcelSaverBase
    {
        internal string excel_path;
        internal string excel_sheetname;

        public ExcelSaverBase(string model) {
            Type setting = typeof(env);
            this.excel_path = (string)setting.GetField("excel_path_" + model).GetValue(null);
            this.excel_sheetname = (string)setting.GetField("excel_sheetname_" + model).GetValue(null);
        }

        public ExcelSaverBase() { }

        public virtual void Excel_Save(Dictionary<string, string> new_item) {}
    }
}
