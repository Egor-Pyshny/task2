using Data.models;
using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic
{
    public class ExcelSaver_iuh : ExcelSaverBase
    {
        public ExcelSaver_iuh() : base("iuh") { }

        public override void Excel_Save(Dictionary<string, string> new_item)
        {
            ExcelMapper excelMapper = new ExcelMapper(excel_path);
            List<iuh_model> iuh = new ExcelMapper(excel_path).Fetch<iuh_model>().ToList();
            iuh_model temp = (iuh_model)new_item;
            for (int i = 0; i < iuh.Count; i++)
            {

                if (iuh[i].ynp == temp.ynp)
                {
                    iuh.RemoveAt(i);
                    break;
                }
            }
            iuh.Add(temp);
            excelMapper.Save(excel_path, iuh, excel_sheetname);
        }
    }
}
