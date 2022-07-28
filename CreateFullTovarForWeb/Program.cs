using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using System.Linq;
using System.Globalization;
using Renci.SshNet;

namespace CreateFullTovarForWeb
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
                DeleteOldFiles(DataDirectory);

                bool Compare = System.Configuration.ConfigurationManager.AppSettings["Compare"] == "1";
                bool useFtp = System.Configuration.ConfigurationManager.AppSettings["useFTP"] == "1";
                bool useSFtp = System.Configuration.ConfigurationManager.AppSettings["useSFTP"] == "1";
                string CompareDirectory = Path.Combine(Environment.CurrentDirectory, System.Configuration.ConfigurationManager.AppSettings["CompareFolder"]);

                string DataFile = DataDirectory + "\\" + DateTime.Now.ToShortDateString().Replace('.', '-') + "-GetFullTovarForWeb.csv";
                string CompareFile = CompareDirectory + "\\Compare-" + DateTime.Now.ToShortDateString().Replace('.', '-') + "-GetFullTovarForWeb.csv";

                bool FilesEqual = false;
                List<DataRow> OldDataList = new List<DataRow>();
                if (Compare)
                {
                    DeleteOldFiles(CompareDirectory);
                    DirectoryInfo Dir = new DirectoryInfo(CompareDirectory);
                    var listCompFiles = Dir.GetFiles("*GetFullTovarForWeb.csv");
                    if (listCompFiles.Count() > 0)
                    {
                        FileInfo LastFile = listCompFiles.OrderByDescending(f => f.LastWriteTime).First();
                        using (StreamReader sr = new StreamReader(LastFile.FullName, Encoding.UTF8))
                        {
                            int num = 0;
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                num++;
                                if (num > 1)
                                {
                                    DataRow row = new DataRow();
                                    row.Data = line.Split(';');
                                    row.Line = line;
                                    row.Code = row.Data[0];

                                    OldDataList.Add(row);
                                }
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
                //Console.WriteLine(DateTime.Now + ": " + "Start");
                try
                {
                    if (!string.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings["OrderFolder"]))
                    {
                        DirectoryInfo DirOrder = new DirectoryInfo(System.Configuration.ConfigurationManager.AppSettings["OrderFolder"]);
                        var files = DirOrder.GetFiles("*.csv");
                        foreach (FileInfo file in files)
                        {
                            Log(file.FullName);
                            Log(webClient.CreateDocuments(File.ReadAllBytes(file.FullName)));
                            file.Delete();
                        }
                    }

                    pr = webClient.GetFullTovarForWeb(false);
                    if (pr.ReplyState.State == "Success")
                    {
                        if (pr.ExcelFile.Length == 0)
                            Log("Возвращенный файл пустой");
                        else
                        {
                            if (Compare)
                            {
                                File.WriteAllBytes(CompareFile, pr.ExcelFile);
                                if (useFtp)
                                    UploadToFTP(new FileInfo(CompareFile), true);
                                if (useSFtp)
                                    UploadToSFTP(new FileInfo(CompareFile), true);
                                Compare = (OldDataList.Count > 0);
                            }
                            if (Compare)
                            {
                                List<DataRow> ReplyData = new List<DataRow>();
                                List<string> NewData = new List<string>();
                                int fNumber = 0;
                                int fRozn = 0;
                                using (Stream ms = new MemoryStream(pr.ExcelFile))
                                {
                                    using (StreamReader sr = new StreamReader(ms, Encoding.UTF8))
                                    {
                                        int num = 0;
                                        string line;
                                        while ((line = sr.ReadLine()) != null)
                                        {
                                            num++;
                                            if (num == 1)
                                            {
                                                NewData.Add(line);
                                                string[] header = line.ToLower().Split(';');
                                                fNumber = header.Count();
                                                fRozn = Array.IndexOf(header, "розничная");
                                            }
                                            else
                                            {
                                                DataRow row = new DataRow();
                                                row.Data = line.Split(';');
                                                row.Line = line;
                                                row.Code = row.Data[0];
                                                ReplyData.Add(row);
                                            }
                                        }

                                    }
                                }
                                foreach (DataRow row in ReplyData)
                                {
                                    DataRow Oldrow = OldDataList.FirstOrDefault(x => x.Code == row.Code);
                                    if (Oldrow != null)
                                    {
                                        Oldrow.Used = 1;
                                        if (!String.Equals(row.Line, Oldrow.Line))
                                            NewData.Add(row.Line);
                                    }
                                    else
                                        NewData.Add(row.Line);
                                }
                                foreach (DataRow row in OldDataList.Where(s => s.Used == 0).ToList())
                                {
                                    if (row.Data.Count() > 11 + fRozn)
                                    {
                                        for (int i = 0; i <= 11; i++)
                                            row.Data[i + fRozn] = "";
                                        string newLine = string.Join(";", row.Data);
                                        if (row.Data.Count() < fNumber)
                                            newLine = newLine + new string(';', (fNumber - row.Data.Count()));
                                        NewData.Add(newLine);
                                    }
                                }
                                if (NewData.Count > 1)
                                    File.WriteAllLines(DataFile, NewData);
                                FilesEqual = (NewData.Count == 0);
                            }
                            else
                                File.WriteAllBytes(DataFile, pr.ExcelFile);
                            if (!File.Exists(DataFile)) // && (UploadToFTP(new FileInfo(DataFile))))
                            {
                                //Log("Удачное завершение");
                                Log("Удачное завершение. Файл пустой.");
                                WorkDone(DateTime.Today);
                            }
                            else if (FilesEqual)
                            {
                                Log("Удачное завершение. Номенклатура не изменилась");
                                WorkDone(DateTime.Today);
                            }
                            else
                            {
                                if ((useFtp && UploadToFTP(new FileInfo(DataFile))) | (useSFtp && UploadToSFTP(new FileInfo(DataFile))))
                                {
                                    Log("Удачное завершение");
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
                //Console.WriteLine(DateTime.Now + ": " + "Ready");
                //Console.ReadKey();
                Log("Ending...");
            }
        }

        static void Log(string logMessage)
        {
            string logFile = Environment.CurrentDirectory + @"\log.txt";
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
                string journal = Environment.CurrentDirectory + @"\journal.txt";
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

        static void DeleteOldFiles(string folder)
        {
            DirectoryInfo Dir = new DirectoryInfo(folder);
            var files = Dir.GetFiles("*GetFullTovarForWeb.csv");
            int maxFiles = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SaveFilesDays"]);
            if (files.Length > maxFiles)
            {
                int passFiles = 0;
                foreach (FileInfo file in files.OrderByDescending(f => f.LastWriteTime).ToList())
                {
                    passFiles++;
                    if (passFiles > maxFiles)
                        file.Delete();
                    //if (DateTime.Now - File.GetCreationTime(file) > TimeSpan.FromDays(6d))
                    //{
                    //    File.Delete(file);
                    //}
                }
            }
        }

        static bool UploadToFTP(FileInfo file, bool IsFull = false)
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["FTP"];
            if (string.IsNullOrEmpty(path))
            {
                Log("Не указан ftp");
                return false;
            }
            FtpStatusCode ftpStatus = FtpStatusCode.Undefined;
            string fileName = "export" + (IsFull ? "Full" : "") + ".csv";
            try
            {
                FtpWebRequest ftpClient = (FtpWebRequest)FtpWebRequest.Create(path + fileName);
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
                    Log(fileName + " : " + uploadResponse.StatusDescription);
                uploadResponse.Close();
            }
            catch (Exception e)
            {
                Log(e.Message);
            }
            Log(fileName + " : " + ftpStatus.ToString());
            //UploadToSFTP(file, IsFull);
            return ftpStatus == FtpStatusCode.ClosingData;
        }

        static bool UploadToSFTP(FileInfo file, bool IsFull = false)
        {
            bool result = false;
            string fileName = "export" + (IsFull ? "Full" : "") + ".csv";
            string sftp_url = System.Configuration.ConfigurationManager.AppSettings["sFTP"];
            string sftp_path = System.Configuration.ConfigurationManager.AppSettings["sftpPath"];
            string sftp_user = System.Configuration.ConfigurationManager.AppSettings["sftpUser"];
            string sftp_password = System.Configuration.ConfigurationManager.AppSettings["sftpPassword"];
            if (!string.IsNullOrEmpty(sftp_url) && !string.IsNullOrEmpty(sftp_path) && !string.IsNullOrEmpty(sftp_user) && !string.IsNullOrEmpty(sftp_password))
            {
                int port = 22;
                int.TryParse(System.Configuration.ConfigurationManager.AppSettings["sftpPort"], out port);
                if (port == 0)
                    port = 22;
                using (SftpClient sftp = new SftpClient(sftp_url, port, sftp_user, sftp_password))
                {
                    try
                    {
                        sftp.Connect();
                        sftp.ChangeDirectory(sftp_path);
                        using (FileStream fs = file.OpenRead())
                        {
                            sftp.UploadFile(fs, fileName, true);
                        }
                        sftp.Disconnect();
                        Log(fileName + " : sFtp finished upload successfully");
                        result = true;
                    }
                    catch (Exception e)
                    {
                        Log(e.Message);
                    }
                }
            }
            return result;
        }

        class DataRow
        {
            public string Code;
            public string Line;
            public string[] Data;
            public int Used;
        }

    }
}
