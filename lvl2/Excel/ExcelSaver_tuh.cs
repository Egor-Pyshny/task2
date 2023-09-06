using Data.models;
using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic
{
    public class ExcelSaver_tuh : ExcelSaverBase
    {
        public ExcelSaver_tuh() : base("tuh") {}

        public override void Excel_Save(Dictionary<string, string> new_item)
        {
            ExcelMapper excelMapper = new ExcelMapper(excel_path);
            List<tuh_model> tuh = new ExcelMapper(excel_path).Fetch<tuh_model>().ToList();
            tuh_model temp = (tuh_model)new_item;
            for (int i = 0; i < tuh.Count; i++)
            {
                if (tuh[i].id == temp.id)
                {
                    tuh.RemoveAt(i);
                    break;
                }
            }
            tuh.Add(temp);
            excelMapper.Save(excel_path, tuh, excel_sheetname);
        }
    }
}
