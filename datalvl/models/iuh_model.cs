using System.Collections.Generic;
using ColumnAttribute = Ganss.Excel.ColumnAttribute;

namespace Data.models
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

        public static explicit operator iuh_model(Dictionary<string, string> dict)
        {
            iuh_model iuh = new iuh_model();
            iuh.full_name = dict["Полное наименование:"];
            iuh.ynp = dict["УНП:"];
            iuh.creation_date = dict["Дата создания:"];
            iuh.day_summ = dict["Сумма в день:"];
            iuh.day_amount = dict["Кол-во в день:"];
            iuh.month_sum = dict["Сумма в месяц:"];
            iuh.month_amount = dict["Кол-во в месяц:"];
            iuh.responsible = "";
            iuh.status = "Работает";
            return iuh;
        }
    }
}
