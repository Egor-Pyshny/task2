using System.Collections.Generic;
using ColumnAttribute = Ganss.Excel.ColumnAttribute;

namespace Data.models
{
    public class tuh_model
    {
        [Column("ID")]
        public string id { get; set; }

        [Column("Лимит сутки")]
        public string day_transaction_limit { get; set; }

        [Column("Кол-во сумма сутки")]
        public string day_sum { get; set; }

        [Column("Лимит месяц")]
        public string month_transaction_limit { get; set; }

        [Column("Кол-во сумма месяц")]
        public string month_sum { get; set; }

        [Column("Точка установки")]
        public string address { get; set; }

        [Column("Дата установки")]
        public string setup_date { get; set; }

        [Column("Лицензия")]
        public string license { get; set; }

        [Column("Дата окончания лицензии")]
        public string valid_date { get; set; }

        public static explicit operator tuh_model(Dictionary<string, string> dict)
        {
            tuh_model tuh = new tuh_model();
            tuh.id = dict["ID:"];
            tuh.day_transaction_limit = dict["Лимит транзакций в сутки:"];
            tuh.day_sum = dict["Кол-во сумма в сутки:"];
            tuh.month_transaction_limit = dict["Лимит транзакций в месяц:"];
            tuh.month_sum = dict["Кол-во сумма в месяц:"];
            tuh.address = dict["Адрес:"];
            tuh.setup_date = dict["Дата установки:"];
            tuh.license = dict["Лицензия:"] != "" ? dict["Лицензия:"] : "НЕТ";
            tuh.valid_date = dict["Действут до:"];
            return tuh;
        }
    }
}
