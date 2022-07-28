using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;

namespace CreateSohSellingDocs
{
    class Program
    {
        static void Main(string[] args)
        {
            string DataDirectory = Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DataFolder"]);
            if (!Directory.Exists(DataDirectory))
                Directory.CreateDirectory(DataDirectory);
            using (Exchange1C.Exchange1C77 WebClient = new Exchange1C.Exchange1C77())
            {
                WebClient.Url = ConfigurationManager.AppSettings["URL"];
                WebClient.Timeout = Convert.ToInt32(ConfigurationManager.AppSettings["Timeout"]); 
                Log("Starting...");
                try
                {
                    DirectoryInfo DirOrder = new DirectoryInfo(DataDirectory);
                    var files = DirOrder.EnumerateFiles("*.csv", SearchOption.AllDirectories);
                    foreach (FileInfo file in files)
                    {
                        Log("Файл " + file.Name);
                        Log(WebClient.CreateDocuments(File.ReadAllBytes(file.FullName)));
                    }
                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
                Log("Ending...");
            }
        }
        static void Log(string logMessage)
        {
            string[] LogMessages = logMessage.Split(new[] { "\r\n","\r","\n"}, StringSplitOptions.None);
            string logFile = Environment.CurrentDirectory + @"\log.txt";
            using (StreamWriter w = File.AppendText(logFile))
            {
                for (int i = 0; i < LogMessages.Length; i++)
                    w.WriteLine("{0} : {1}", DateTime.Now.ToString(), LogMessages[i]);
            }
        }
    }
}
