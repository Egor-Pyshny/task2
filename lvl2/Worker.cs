using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Data;
using AutoIt;

namespace Logic
{
    public class Worker
    {
        public async Task StartWorkAsync(){
            Setup();
            Timer timer = new Timer(DoWork, null, 0, env.timer_period_minutes * 60000);
        }

        private void Setup()
        {
            var files = Directory.GetFiles(env.project_dirrectory);
            if (!File.Exists(env.excel_path_iuh)) File.Create(env.excel_path_iuh).Close();
            if (!File.Exists(env.excel_path_tuh)) File.Create(env.excel_path_tuh).Close();
            if (!Directory.Exists(env.error_directory)) Directory.CreateDirectory(env.error_directory);
        }

        private void DoWork(object st) {
            if (!Directory.Exists(env.working_directory_path)) return;
            AutoItX.WinClose($"{env.today_date}");
            Dictionary<string, int> correct_files = new Dictionary<string, int>();
            try
            {
                string formats = "";
                foreach (string str in env.possible_models.Keys) {
                    formats += ("|" + str);
                }
                formats = formats.Remove(0, 1);
                var dir = new DirectoryInfo(env.working_directory_path);
                AutoItX.Run($"explorer.exe {env.working_directory_path}","");
                int i = 1;
                foreach (FileInfo f in dir.GetFiles()) {
                    if (Regex.IsMatch(f.Name, $@"^РЕГИСТРАЦИЯ.*(?:{formats}){f.Extension}$")) {
                        if (check_files(f))
                        {
                            correct_files.Add(f.Name, i);
                        }
                    }
                    i++;
                }             
            }
            catch(Exception e){ }
            int prev_pos = 0;
            foreach (KeyValuePair<string,int> pair in correct_files) {
                parse_file(pair, prev_pos);
                prev_pos = pair.Value;
            }
            AutoItX.WinWait(env.today_date);
            AutoItX.WinClose($"{env.today_date}");
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

        private void parse_file(KeyValuePair<string, int> filename, int prev_pos)
        {
            Type setting = typeof(env);
            ExcelSaverBase excelSaver;
            string type_mod = null;
            foreach (KeyValuePair<string, string> pair in env.possible_models) {
                if (filename.Key.Contains(pair.Key)) {
                    type_mod = pair.Value;
                    break;
                }
            }
            string[] fields = (string[])setting.GetField("fields_" + type_mod).GetValue(null);
            string[] possible_endings = (string[])setting.GetField("possible_endings_" + type_mod).GetValue(null);
            string necessary_field = (string)setting.GetField("necessary_field_" + type_mod).GetValue(null);
            Type saver_type = Type.GetType("Logic.ExcelSaver_" + type_mod);
            excelSaver = (ExcelSaverBase) Activator.CreateInstance(saver_type);
            string path = env.working_directory_path + "\\" + filename.Key;            
            string context = open_and_read(filename, prev_pos);
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
                else matches.Add(field ,match);
            }
            if (matches[necessary_field] != "")
            {
                excelSaver.Excel_Save(matches);                    
            }
            else
            {
                try
                {
                    File.Copy(env.working_directory_path + $"\\{filename.Key}", env.error_directory + $"\\{filename.Key}");
                }
                catch (System.IO.IOException) { }
            }
        }

        private string open_and_read(KeyValuePair<string, int> pair,int prev_pos)
        {
            AutoItX.ClipPut("");
            AutoItX.WinWait(env.today_date);
            for (int i = 0; i < prev_pos; i++)
            {
                AutoItX.Send("{UP}");
            }
            for (int i = 0; i < pair.Value; i++)
            {
                AutoItX.Send("{DOWN}");
            }
            AutoItX.Send("{UP}");
            AutoItX.Send("{ENTER}");            
            Thread.Sleep(500);
            var title = AutoItX.WinGetTitle("[ACTIVE]");
            AutoItX.Send("^a");
            AutoItX.Send("^c");
            string buffer = "";
            while (buffer == "") { buffer = AutoItX.ClipGet(); }
            AutoItX.WinClose(title);
            return buffer;
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
