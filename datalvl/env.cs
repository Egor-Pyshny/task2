using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace datalvl
{
    public static class env
    {
        private static string today_date = DateTime.Now.ToString().Split(' ')[0];
        private static string yesterday = DateTime.Now.AddDays(-1).ToString("dd.MM.yyyy");
        public static string working_directory_path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\"+today_date;
        public static string excel_path_iuh = working_directory_path + "\\ИЭ.xlsx";
        public static string excel_path_tuh = working_directory_path + "\\ТЭ.xlsx";
        public static string excel_sheetname_iuh = "Лист1";
        public static string excel_sheetname_tuh = "Лист1";
        public static string necessary_field_iuh = "УНП";
        public static string delimiter = "☢";
        public static string project_dirrectory = Environment.CurrentDirectory.ToString();
        public static string error_directory = project_dirrectory + "\\Errors_" +today_date;
        public static string check_file = project_dirrectory + "\\first_run_" + today_date + ".txt";
        public static string yesterday_check_file = project_dirrectory + "||first_run_" + yesterday + ".txt";
        public static string[] data_fields = {
            "Дата установки:",
            "Дата создания:",
            "Действут до:",
        };
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
        public static List<FileInfo> files = new List<FileInfo>();
        public static string[] fields_iuh = {
            "Полное наименование:",
            "УНП:",
            "Сумма в день:",
            "Дата создания:",
            "Сумма в месяц:",
            "Кол-во в день:",
            "Кол-во в месяц:",
        };
        public static string[] fields_tuh = {
            "ID",
            "Лимит транзакций в сутки:",
            "Кол-во сумма в сутки:",
            "Лимит транзакций в месяц:",
            "Кол-во сумма в месяц:",
            "Адрес:",
            "Дата установки:",
            "Лицензия:",
            "Действут до:"
        };
        public static int timer_period_minutes = 30;
    }
}
