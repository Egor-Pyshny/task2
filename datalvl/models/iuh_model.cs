using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using Ganss.Excel;
using System.Text;
using System.Threading.Tasks;
using ColumnAttribute = Ganss.Excel.ColumnAttribute;

namespace datalvl.models
{
    public class iuh_model
    {
        [Column("Наименование")]
        public string full_name { get; set; }

        [Column("УНП")]
        public string ynp { get; set; }

        [Column("Дата создания")]
        public string creation_date { get; set; }

        [Column("Сумма в день")]
        public string day_summ { get; set; }

        [Column("Кол-во в день")]
        public string day_amount { get; set; }

        [Column("Сумма в месяц")]
        public string month_sum { get; set; }

        [Column("Кол-во в месяц")]
        public string month_amount { get; set; }

        [Column("Ответственный")]
        public string responsible { get; set; }

        [Column("Статус")]
        public string status { get; set; } = "Работает";

        public static explicit operator iuh_model(List<string> v)
        {
            iuh_model iuh = new iuh_model();
            iuh.full_name = v[0];
            iuh.ynp = v[1];
            iuh.creation_date = v[3];
            iuh.day_summ = v[2];
            iuh.day_amount = v[5];
            iuh.month_sum = v[4];
            iuh.month_amount = v[6];
            iuh.responsible = "";
            iuh.status = "Работает";
            return iuh;
        }
    }
}
