using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Globalization;

namespace CreateProdagiTovara
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!WorkDone(DateTime.Today, true))
            {
                string DataDirectory = Path.Combine(Environment.CurrentDirectory, System.Configuration.ConfigurationManager.AppSettings["DataFolder"]);
                if (!Directory.Exists(DataDirectory))
                    Directory.CreateDirectory(DataDirectory);
                DeleteOldFiles(DataDirectory, "*DailySales.csv");

                bool Compare = System.Configuration.ConfigurationManager.AppSettings["Compare"] == "1";
                string CompareDirectory = Path.Combine(Environment.CurrentDirectory, System.Configuration.ConfigurationManager.AppSettings["CompareFolder"]);

                string DataFile = DataDirectory + "\\" + DateTime.Now.ToShortDateString().Replace('.', '-') + "-DailySales.csv";
                string CompareFile = CompareDirectory + "\\" + DateTime.Now.ToShortDateString().Replace('.', '-') + "-AllSales.csv";

                bool FilesEqual = false;
                List<DataRow> OldDataList = new List<DataRow>();
                if (Compare)
                {
                    DeleteOldFiles(CompareDirectory, "*AllSales.csv");
                    DirectoryInfo Dir = new DirectoryInfo(CompareDirectory);
                    var listCompFiles = Dir.GetFiles("*AllSales.csv");
                    if (listCompFiles.Count() > 0)
                    {
                        FileInfo LastFile = listCompFiles.OrderByDescending(f => f.LastWriteTime).First();
                        using (StreamReader sr = new StreamReader(LastFile.FullName, Encoding.UTF8))
                        {
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                DataRow row = new DataRow();
                                row.Data = line.Split(';');
                                row.Line = line;
                                row.CustCode = row.Data[0];
                                row.Code = row.Data[1];

                                OldDataList.Add(row);
                            }
                        }
                    }
                }

                Exchange1C.Exchange1C77 webClient = new Exchange1C.Exchange1C77();
                webClient.Url = System.Configuration.ConfigurationManager.AppSettings["URL"];
                webClient.Timeout = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Timeout"]); //360000; 

                //webClient.Proxy = new WebProxy("127.0.0.1", 8888);
                Exchange1C.PriceList pr = new Exchange1C.PriceList();
                Log("Starting...");
                try
                {
                    pr = webClient.GetFullTovarProdagiForWeb();
                    if (pr.ReplyState.State == "Success")
                    {
                        if (pr.ExcelFile.Length == 0)
                            Log("Возвращенный файл пустой");
                        else
                        {
                            if (Compare)
                            {
                                File.WriteAllBytes(CompareFile, pr.ExcelFile);
                                UploadToFTP(new FileInfo(CompareFile));
                                Compare = (OldDataList.Count > 0);
                            }
                            if (Compare)
                            {
                                List<DataRow> ReplyData = new List<DataRow>();
                                List<string> NewData = new List<string>();
                                using (Stream ms = new MemoryStream(pr.ExcelFile))
                                {
                                    using (StreamReader sr = new StreamReader(ms, Encoding.UTF8))
                                    {
                                        string line;
                                        while ((line = sr.ReadLine()) != null)
                                        {
                                            DataRow row = new DataRow();
                                            row.Data = line.Split(';');
                                            row.Line = line;
                                            row.CustCode = row.Data[0];
                                            row.Code = row.Data[1];
                                            ReplyData.Add(row);
                                        }
                                    }
                                }
                                foreach (DataRow row in ReplyData)
                                {
                                    DataRow Oldrow = OldDataList.FirstOrDefault(x => x.CustCode == row.CustCode && x.Code == row.Code);
                                    if (Oldrow == null)
                                        NewData.Add(row.Line);
                                }
                                if (NewData.Count > 1)
                                    File.WriteAllLines(DataFile, NewData);
                                FilesEqual = (NewData.Count == 0);
                            }
                            else
                                File.WriteAllBytes(DataFile, pr.ExcelFile);


                            if ((File.Exists(DataFile)) && (UploadToFTP(new FileInfo(DataFile))))
                            {
                                Log("Удачное завершение");
                                WorkDone(DateTime.Today);
                            }
                            else if (FilesEqual)
                            {
                                Log("Удачное завершение. Продаж не было");
                                WorkDone(DateTime.Today);
                            }
                            else
                            {
                                Log("Удаляем последние сформированные файлы из-за ошибки ftp");
                                if (File.Exists(DataFile))
                                    File.Delete(DataFile);
                                if (File.Exists(CompareFile))
                                    File.Delete(CompareFile);
                            }
                        }
                    }
                    else
                        Log(pr.ReplyState.Description);

                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
                finally
                {
                    webClient.Dispose();
                }
                Log("Ending...");
            }
        }

        static void Log(string logMessage)
        {
            string logFile = Environment.CurrentDirectory + @"\logProdagi.txt";
            using (StreamWriter w = File.AppendText(logFile))
            {
                w.WriteLine("{0} : {1}", DateTime.Now.ToString(), logMessage);
            }
        }

        static bool WorkDone(DateTime CurrentDate, bool Check = false)
        {
            bool workDone = false;
            try
            {
                string journal = Environment.CurrentDirectory + @"\journalProdagi.txt";
                if (!File.Exists(journal))
                    File.Create(journal).Close();
                if (Check)
                {
                    DateTime ReadDate;
                    using (StreamReader r = new StreamReader(journal))
                    {
                        while (r.Peek() >= 0)
                        {
                            DateTime.TryParseExact(r.ReadLine(), "d", CultureInfo.InvariantCulture, DateTimeStyles.None, out ReadDate);
                            workDone = (ReadDate == CurrentDate);
                            if (workDone)
                                break;
                        }
                    }
                }
                else
                {
                    using (StreamWriter w = File.AppendText(journal))
                    {
                        w.WriteLine(CurrentDate.ToString("d", CultureInfo.InvariantCulture));
                        workDone = true;
                    }

                }
                return workDone;
            }
            catch (Exception e)
            {
                Log(e.Message);
                return workDone;
            }
        }

        static void DeleteOldFiles(string folder, string pattern)
        {
            DirectoryInfo Dir = new DirectoryInfo(folder);
            var files = Dir.GetFiles(pattern);
            int maxFiles = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SaveFilesDays"]);
            if (files.Length > maxFiles)
            {
                int passFiles = 0;
                foreach (FileInfo file in files.OrderByDescending(f => f.LastWriteTime).ToList())
                {
                    passFiles++;
                    if (passFiles > maxFiles)
                        file.Delete();
                }
            }
        }

        static bool UploadToFTP(FileInfo file, string ftpFileName = "")
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["FTP"];
            if (string.IsNullOrEmpty(path))
            {
                Log("Не указан ftp");
                return false;
            }
            if (string.IsNullOrEmpty(ftpFileName))
            {
                ftpFileName = file.Name;
            }
            FtpStatusCode ftpStatus = FtpStatusCode.Undefined;
            try
            {
                FtpWebRequest ftpClient = (FtpWebRequest)FtpWebRequest.Create(path + ftpFileName);
                string ftp_user = System.Configuration.ConfigurationManager.AppSettings["ftpUser"];
                string ftp_password = System.Configuration.ConfigurationManager.AppSettings["ftpPassword"];
                if (string.IsNullOrEmpty(ftp_user))
                    ftp_user = "anonymous";
                if ((ftp_user == "anonymous") && (string.IsNullOrEmpty(ftp_password)))
                    ftp_user = "m@mm.ru";
                ftpClient.Credentials = new System.Net.NetworkCredential(ftp_user, ftp_password);
                ftpClient.Method = WebRequestMethods.Ftp.UploadFile;
                ftpClient.KeepAlive = true;
                ftpClient.UseBinary = System.Configuration.ConfigurationManager.AppSettings["Binary"] == "1";
                ftpClient.UsePassive = System.Configuration.ConfigurationManager.AppSettings["Passive"] == "1";
                ftpClient.ContentLength = file.Length;
                byte[] buffer = new byte[4097];
                int bytes = 0;
                int total_bytes = (int)file.Length;
                using (FileStream fs = file.OpenRead())
                {
                    using (Stream rs = ftpClient.GetRequestStream())
                    {
                        while (total_bytes > 0)
                        {
                            bytes = fs.Read(buffer, 0, buffer.Length);
                            rs.Write(buffer, 0, bytes);
                            total_bytes = total_bytes - bytes;
                        }
                    }
                }
                FtpWebResponse uploadResponse = (FtpWebResponse)ftpClient.GetResponse();
                ftpStatus = uploadResponse.StatusCode;
                if (ftpStatus != FtpStatusCode.ClosingData)
                    Log(uploadResponse.StatusDescription);
                uploadResponse.Close();
            }
            catch (Exception e)
            {
                Log(e.Message);
            }
            return (ftpStatus == FtpStatusCode.ClosingData);
        }

        class DataRow
        {
            public string CustCode;
            public string Code;
            public string Line;
            public string[] Data;
            public int Used;
        }
    }
}
