using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using datalvl.models;

namespace lvl2
{
    public class Worker
    {
        public async void StartWorkAsync(){
            Timer timer = new Timer(DoWork, null, 0, 30 * 60000);
        }

        public void DoWork(object st)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string date = DateTime.Now.ToString().Split(' ')[0];
            string path = desktop + $"\\{date}";
            List<string> correct_files = new List<string>();
            try
            {
                var dir = new DirectoryInfo(path);
                foreach (FileInfo f in dir.GetFiles()) {
                    if (Regex.IsMatch(f.Name, $@"^РЕГИСТРАЦИЯ.*(?:ИЭ|ТЭ){f.Extension}$")) {
                        correct_files.Add(f.Name);
                    }
                }             
            }
            catch(Exception e){ }
            foreach (string filename in correct_files) {
                if (filename.Contains("ИЭ")) parse_iuh(path+"\\"+filename); else parse_tuh();
            }
        }

        public void parse_iuh(string path){
            using (StreamReader reader = new StreamReader(path)){
                var context = reader.ReadToEnd();
                string[] tmp = new string[] {
                    "Полное наименование:",
                    "УНП:",
                    "Сумма в день:",
                    "Дата создания:",
                    "Сумма в месяц:",
                    "Кол-во в день:",
                };
                string[] patterns = new string[] {
                    $@"{tmp[0]}[^.]*.",
                    $@"{tmp[1]}.*",
                    $@"{tmp[2]}.*",
                    $@"{tmp[3]}.*",
                    $@"{tmp[4]}.*",
                    $@"{tmp[5]}.*",
                };
                List<string> matches = new List<string>();
                int l = 0;
                foreach (string pattern in patterns) {
                    matches.Add(Regex.Match(context, pattern).Value.Remove(0,tmp[l].Length).Trim());
                    l++;
                }
                matches[3] = parse_data(matches[3]);
                iuh_model iuh = new iuh_model(matches.ToArray());
                ExcelMapper mapper = new ExcelMapper("ИЭ.xlsx");


            }
        }

        private string parse_data(string data)
        {
            var tmp = data.Split(' ');
            Dictionary<string, string> dict = new Dictionary<string, string>()
            {
                { "января" ,"01"},
                { "февраля" ,"02"},
                { "марта" ,"03"},
                { "апреля" ,"04"},
                { "мая" ,"05"},
                { "июня" ,"06"},
                { "июля" ,"07"},
                { "августа" ,"08"},
                { "сентября" ,"09"},
                { "октября" ,"10"},
                { "ноября" ,"11"},
                { "декабря" ,"12"}
            };
            return tmp[0] + "." + dict[tmp[1]] + "." + tmp[2] + " " + tmp[4];
        }

        public void parse_tuh() { }
    }
}
