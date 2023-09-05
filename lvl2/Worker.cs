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
using datalvl;

namespace lvl2
{
    public class Worker
    {
        public async void StartWorkAsync(){
            Setup();
            Timer timer = new Timer(DoWork, null, 0, 30 * 60000);
        }

        private void Setup()
        {
            var files = Directory.GetFiles(env.project_dirrectory);
            if (!files.Contains(env.check_file)) {
                Directory.CreateDirectory(env.error_directory);
                File.Create(env.check_file);
                if (files.Contains(env.yesterday_check_file)) {
                    File.Delete(env.yesterday_check_file);
                }
            }
        }

        public void DoWork(object st) {

            List<string> correct_files = new List<string>();
            try
            {
                var dir = new DirectoryInfo(env.working_directory_path);
                foreach (FileInfo f in dir.GetFiles()) {
                    if (Regex.IsMatch(f.Name, $@"^РЕГИСТРАЦИЯ.*(?:ИЭ|ТЭ){f.Extension}$")) {
                        if (check_files(f))
                        {
                            correct_files.Add(f.Name);
                            env.files.Add(f);
                        }
                    }
                }             
            }
            catch(Exception e){ }

            foreach (string filename in correct_files) {
                if (filename.Contains("ИЭ")) parse_iuh(filename); else parse_tuh();
            }
        }

        private bool check_files(FileInfo info)
        {
            bool res;
            DateTime modified = info.LastWriteTime;
            string name = info.Name;

            return res;
        }

        public void parse_iuh(string filename)
        {
            string path = env.working_directory_path + "\\" + filename;
            using (StreamReader reader = new StreamReader(path))
            {
                var context = reader.ReadToEnd();
                List<string> matches = new List<string>();
                foreach (string field in env.fields_iuh)
                {
                    string tmp = context;
                    foreach (string ending in env.possible_endings_iuh)
                    {
                        if (!field.Contains(ending)) tmp = tmp.Replace(ending, env.delimiter);
                    }
                    string pattern = $@"{field}[^{env.delimiter}]*";
                    string match = Regex.Match(tmp, pattern).Value;
                    if (match != "")
                    {
                        matches.Add(field == "Дата создания:" ? parse_data(match.Remove(0, field.Length).Trim()) : match.Remove(0, field.Length).Trim());
                    }
                    else if (field == "УНП")
                    {
                        break;
                    }
                    else matches.Add(match);
                }
                /*string[] patterns = new string[] {
                    $@"{fields[0]}[^.]*",
                };*/
                if (matches[1] != "")
                {
                    ExcelMapper excelMapper = new ExcelMapper(env.excel_iuh_path);
                    List<iuh_model> iuh = new ExcelMapper(env.excel_iuh_path).Fetch<iuh_model>().ToList();
                    iuh_model temp = (iuh_model)matches;
                    for (int i = 0; i < iuh.Count; i++)
                    {
                        if (iuh[i].ynp == temp.ynp)
                        {
                            iuh.RemoveAt(i);
                            break;
                        }
                    }
                    iuh.Add(temp);
                    excelMapper.Save(env.excel_iuh_path, iuh, env.excel_iuh_sheetname);
                }
                else
                {
                    //проверить все диретории и создать если нет для ошибок
                    //записать куда то там
                }
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
            return tmp[0] + "." + dict[tmp[1]] + "." + tmp[2] + " " + ((tmp.Length>4)?(tmp[4]):(""));
        }

        public void parse_tuh() { }
    }
}
