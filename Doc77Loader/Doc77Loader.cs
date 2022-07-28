using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.IO;
using System.Configuration;
using Ionic.Zip;

namespace Doc77Loader
{
    public partial class Doc77Loader : ServiceBase
    {
        public Doc77Loader()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Log("Служба запущена");
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 seconds
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }

        protected override void OnStop()
        {
            Log("Служба остановлена");
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            //Log("Tick");
            CreateDocuments();
        }

        private void CreateDocuments()
        {
            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["OrderFolder"]))
            {
                Exchange1C.Exchange1C77 webClient = new Exchange1C.Exchange1C77();
                webClient.Url = ConfigurationManager.AppSettings["URL"];
                webClient.Timeout = Convert.ToInt32(ConfigurationManager.AppSettings["Timeout"]); //360000; 
                try
                {
                    DirectoryInfo DirOrder = new DirectoryInfo(ConfigurationManager.AppSettings["OrderFolder"]);
                    var files = DirOrder.EnumerateFiles("*.csv", SearchOption.AllDirectories); 
                    foreach (FileInfo file in files)
                    {
                        Log("Файл " + file.Name);
                        Log(webClient.CreateDocuments(File.ReadAllBytes(file.FullName)));
                        file.Attributes = FileAttributes.Normal;
                        file.Delete();
                    }
                    files = DirOrder.EnumerateFiles("*.*", SearchOption.AllDirectories).Where(s => s.Extension.EndsWith("zip") || s.Extension.EndsWith("rar"));
                    foreach (FileInfo file in files)
                    {
                        Log("Файл " + file.Name);
                        Log(webClient.CreateDocuments(GluingFiles(file)));
                        file.Attributes = FileAttributes.Normal;
                        file.Delete();
                    }
                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
                finally
                {
                    webClient.Dispose();
                }
            }

        }

        private byte[] GluingFiles(FileInfo file)
        {
            List<string> lines = null;
            using (ZipFile archive = ZipFile.Read(file.FullName))
            {
                foreach (ZipEntry entry in archive.Entries)
                {
                    if (entry.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        using (StreamReader sr = new StreamReader(entry.OpenReader(), Encoding.UTF8))
                        {
                            string[] context = sr.ReadToEnd().Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                            if (lines == null)
                                lines = new List<string>(context);
                            else
                                lines.AddRange(context);
                        }
                    }
                }
            }
            if (lines.Count > 0)
            {
                lines = lines.Select(s =>
                {
                    string[] data = s.Split(';');
                    if (data.Length > 5)
                        data[0] += data[0];
                    s = string.Join(";", data);
                    return s;
                }).ToList();
                return Encoding.UTF8.GetBytes(string.Join(Environment.NewLine, lines));
            }
            else
                return null;
        }
        
        private void Log(string logMessage)
        {
            string[] lines = logMessage.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            string logFile = ConfigurationManager.AppSettings["OrderFolder"] + @"\log.txt";
            using (StreamWriter w = File.AppendText(logFile))
            {
                int num = 0;
                foreach (string logtext in lines)
                {
                    num++;
                    if (num == 1)
                        w.WriteLine("{0} : {1}", DateTime.Now.ToString(), logtext);
                    else
                        w.WriteLine("                      {0}", logtext);
                }
            }
        }

    }
}
