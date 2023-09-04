using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace datalvl
{
    public static class env
    {
        public static string working_directory_path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\"+DateTime.Now.ToString().Split(' ')[0];
        public static string excel_iuh_path = working_directory_path + "\\ИЭ.xlsx";
        public static string excel_tuh_path = working_directory_path + "\\ЕЭ.xlsx";
        public static string excel_iuh_sheetname = "Лист1";
        public static string excel_tuh_sheetname = "Лист1";
        public static string delimiter = "☢";
        public static string[] possible_endings_iuh = {
            "Наименование",
            "УНП",
            "Сумма",
            "Дата",
            "Кол-во",
            "Срок",
            "База",
            "MERCHANTID",
            "Полное",
            "ТерминалID"
        };
        public static string[] fields_iuh = {
            "Полное наименование:",
            "УНП:",
            "Сумма в день:",
            "Дата создания:",
            "Сумма в месяц:",
            "Кол-во в день:",
            "Кол-во в месяц:",
        };
        public static string[] possible_endings_tuh = {
            
        };
    }
}
