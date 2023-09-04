using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace datalvl.models
{
    public class iuh_model
    {
        public string full_name { get; set; }
        public string ynp { get; set; }
        public string creation_date { get; set; }
        public string day_summ { get; set; }
        public string day_amount { get; set; }
        public string month_sum { get; set; }
        public string responsible { get; set; }
        public string status { get; set; }

        public iuh_model(params string[] param) {
            this.full_name = param[0];
            this.ynp = param[1];
            this.creation_date = param[3];
            this.day_summ = param[2];
            this.day_amount = param[5];
            this.month_sum = param[4];
            this.responsible = "";
            this.status = "Работает";
        }
    }
}
