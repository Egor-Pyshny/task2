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
            Timer timer = new Timer(DoWork, null, 0, env.timer_period_minutes * 60000);
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
            Console.WriteLine(DateTime.Now.TimeOfDay.ToString());
            List<string> correct_files = new List<string>();
            try
            {
                var dir = new DirectoryInfo(env.working_directory_path);
                foreach (FileInfo f in dir.GetFiles()) {
                    if (Regex.IsMatch(f.Name, $@"^РЕГИСТРАЦИЯ.*(?:ИЭ|ТЭ){f.Extension}$")) {
                        if (check_files(f))
                        {
                            correct_files.Add(f.Name);
                        }
                    }
                }             
            }
            catch(Exception e){ }
            foreach (string filename in correct_files) {
                parse_file(filename);
            }
        }

        private bool check_files(FileInfo info)
        {
            DateTime modified = info.LastWriteTime;
            string name = info.Name;
            foreach (FileInfo fileInfo in env.files) {
                if (fileInfo.Name == name) {
                    if (fileInfo.LastWriteTime == modified)
                    {
                        return false;
                    }
                    else if (fileInfo.LastWriteTime > modified) {
                        env.files.Remove(fileInfo);
                        env.files.Add(info);
                        return true;
                    }
                }
            }
            env.files.Add(info);
            return true;
        }

        public void parse_file(string filename)
        {
            Type model_type = null;
            Type setting = typeof(env);
            string type_mod = null;
            if (filename.Contains("ИЭ"))
            {
                model_type = typeof(iuh_model);
                type_mod = "iuh";
            }
            else if (filename.Contains("ТЭ"))
            {
                model_type = typeof(tuh_model);
                type_mod = "tuh";
            }
            string[] fields = (string[])setting.GetField("fields_" + type_mod).GetValue(null);
            string[] possible_endings = (string[])setting.GetField("possible_endings_" + type_mod).GetValue(null);
            string excel_path = (string)setting.GetField("excel_path_" + type_mod).GetValue(null);
            string excel_sheetname = (string)setting.GetField("excel_sheetname_" + type_mod).GetValue(null);
            string necessary_field = (string)setting.GetField("necessary_field_" + type_mod).GetValue(null);
            string path = env.working_directory_path + "\\" + filename;
            using (StreamReader reader = new StreamReader(path))
            {
                var context = reader.ReadToEnd();
                Dictionary<string , string> matches = new Dictionary<string, string>();
                foreach (string field in fields)
                {
                    string tmp = context;
                    foreach (string ending in possible_endings)
                    {
                        if (!field.Contains(ending)) tmp = tmp.Replace(ending, env.delimiter);
                    }
                    string pattern = $@"{field}[^{env.delimiter}]*";
                    string match = Regex.Match(tmp, pattern).Value;
                    if (match != "")
                    {
                        matches.Add(field ,env.data_fields.Contains(field) ? parse_data(match.Remove(0, field.Length).Trim()) : match.Remove(0, field.Length).Trim());
                    }
                    else if (field == necessary_field)
                    {
                        break;
                    }
                    else matches.Add(field ,match);
                }
                if (matches[necessary_field] != "")
                {
                    ExcelMapper excelMapper = new ExcelMapper(excel_path);
                    
                    var iuh = new ExcelMapper(excel_path).Fetch().ToList();
                    object temp = Convert.ChangeType(matches,model_type);
                    for (int i = 0; i < iuh.Count; i++)
                    {
                        //разобраться с записью в таблицу и чтением
                        /*var dsad = Convert.ChangeType(iuh[i],model_type);
                        if (iuh[i].ynp == temp.ynp)
                        {
                            iuh.RemoveAt(i);
                            break;
                        }*/
                    }
                    iuh.Add(temp);
                    excelMapper.Save(excel_path, iuh, excel_sheetname);
                }
                else
                {
                    using (var stream = File.Create(env.project_dirrectory + $"\\{filename}")) {
                        byte[] data = Encoding.Default.GetBytes(context);
                        stream.Write(data,0,data.Length);
                    }

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
    }
}
