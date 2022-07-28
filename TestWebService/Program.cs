using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;
using Ionic.Zip;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security;
//using System.Runtime;

namespace TestWebService
{
    class Program
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpFile;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }

        public enum ShowCommands : int
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }

        [Flags]
        public enum ShellExecuteMaskFlags : uint
        {
            SEE_MASK_DEFAULT = 0x00000000,
            SEE_MASK_CLASSNAME = 0x00000001,
            SEE_MASK_CLASSKEY = 0x00000003,
            SEE_MASK_IDLIST = 0x00000004,
            SEE_MASK_INVOKEIDLIST = 0x0000000c,   // Note SEE_MASK_INVOKEIDLIST(0xC) implies SEE_MASK_IDLIST(0x04)
            SEE_MASK_HOTKEY = 0x00000020,
            SEE_MASK_NOCLOSEPROCESS = 0x00000040,
            SEE_MASK_CONNECTNETDRV = 0x00000080,
            SEE_MASK_NOASYNC = 0x00000100,
            SEE_MASK_FLAG_DDEWAIT = SEE_MASK_NOASYNC,
            SEE_MASK_DOENVSUBST = 0x00000200,
            SEE_MASK_FLAG_NO_UI = 0x00000400,
            SEE_MASK_UNICODE = 0x00004000,
            SEE_MASK_NO_CONSOLE = 0x00008000,
            SEE_MASK_ASYNCOK = 0x00100000,
            SEE_MASK_HMONITOR = 0x00200000,
            SEE_MASK_NOZONECHECKS = 0x00800000,
            SEE_MASK_NOQUERYCLASSSTORE = 0x01000000,
            SEE_MASK_WAITFORINPUTIDLE = 0x02000000,
            SEE_MASK_FLAG_LOG_USAGE = 0x04000000,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct NativeMessage
        {
            public IntPtr handle;
            public uint msg;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point p;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern UInt32 WaitForSingleObject(IntPtr hHandle, UInt32 dwMilliseconds);

        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //static extern bool PeekMessage(out NativeMessage lpMsg, HandleRef hWnd, uint wMsgFilterMin,
        //   uint wMsgFilterMax, uint wRemoveMsg);

        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("User32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool PeekMessage(out NativeMessage message, IntPtr handle, uint filterMin, uint filterMax, uint flags);

        [DllImport("user32.dll")]
        static extern bool TranslateMessage([In] ref NativeMessage lpMsg);

        [DllImport("user32.dll")]
        static extern IntPtr DispatchMessage([In] ref NativeMessage lpmsg);

        [DllImport("kernel32", SetLastError = true)]
        static extern bool CloseHandle(
                IntPtr hObject   // handle to object
                );

        public const UInt32 INFINITE = 0xFFFFFFFF;
        public const UInt32 WAIT_ABANDONED = 0x00000080;
        public const UInt32 WAIT_OBJECT_0 = 0x00000000;
        public const UInt32 WAIT_TIMEOUT = 0x00000102;
        public const uint SW_NORMAL = 1;
        public static void OpenAs()
        {
            SHELLEXECUTEINFO sei = new SHELLEXECUTEINFO();
            sei.cbSize = Marshal.SizeOf(sei);
            sei.fMask = (uint)ShellExecuteMaskFlags.SEE_MASK_NOCLOSEPROCESS;
            sei.hwnd = (IntPtr)0; 
            //sei.lpVerb = "openas";
            sei.lpFile = @"F:\tmp\Firebird_1_5\bin\gbak.exe";
            sei.lpParameters = @"-c F:\tmp\4.48.5.1\CD\Backup\gfs.fbk localhost:F:\tmp\4.48.5.1\GFS.FDB -user SYSDBA -pas masterkey -v";
            sei.nShow = 1;
            if (ShellExecuteEx(ref sei))
            {
                while (WaitForSingleObject(sei.hProcess, 100) == WAIT_TIMEOUT)
                {
                    
                    NativeMessage Msg;
                    IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                    while (PeekMessage(out Msg, handle, 0, 0, 1))
                    {
                        TranslateMessage(ref Msg);
                        DispatchMessage(ref Msg);
                    }
                }
                CloseHandle(sei.hProcess);
            }
            else
                throw new System.ComponentModel.Win32Exception();
        }
        static void Main(string[] args)
        {
            string columnЗаказСамара = CellReference.ConvertNumToColString(9);
            string cellText = "{ПечАртикул} {ПечНоменклатура}";
            var variants = Regex.Matches(cellText, @"(\A|\s*)\{(\w+)\}(\s+|\z)")
                                    .Cast<Match>()
                                    .Select(m => m.Value.Trim());
            string Header = "Код;Название;Краткое описание;Производитель;Артикул;Мощность, Вт.;Мощность, кВт.;Мощность, л.с.;Напряжение, В.;Напряжение аккумулятора, В.;";
            Header = Header + "Емкость аккумулятора, А.ч.;Тип аккумулятора;Гарантия в месяцах;Вес нетто, кг.;Страна происхождения;Раздел 1;Раздел 2; Раздел 3; Остаток;";
            Header = Header + "Розничная;Розничная спец;Оптовая;Оптовая спец;Валюта 1;Валюта 2;Валюта 3;Валюта 4;";
            Header = Header + "Мин розничная цена;Мин оптовая цена;Размер скидки розн;Размер скидки опт;";
            Header = Header + "Вид рукоятки;Выборка четверти;Высота обработки, мм.;Высота подъема жидкости, м.;Глубина обработки, мм.;Глубина погружения, м.;";
            Header = Header + "Глубина забора жидкости, м.;Дальность выброса снега, м.;Диаметр бура, мм.;Диаметр диска, мм.;Диаметр корпуса, мм.;Диаметр патрона внутренний, мм.;";
            Header = Header + "Диаметр посадочный, мм;Диаметр присоединения, дюйм.;Диаметр сверления, мм.;Диаметр цанги, мм.;Диаметр электродов, мм.;Длина сетевого шнура, м.;";
            Header = Header + "Длина шины в комплекте, дюйм;ЖК-дисплей;Запуск двигателя;Количество дисков;Количество звеньев цепи;Количество скоростей;Комплектация;Крутящий момент, Н.м.;";
            Header = Header + "Лазерная направляющая;Легкий старт;Материал корпуса;Маятниковый ход;Модель двигателя;Объем барабана, л.;Объем гидроаккумулятора, л.;Объем контейнера, л.;";
            Header = Header + "Объем двигателя, см3.;Объем ресивера, л.;Объем топливного бака, л.;Объем травосборника, л.;Патрубок пылеотвода;ПВ на максимальном токе, %;Передача;";
            Header = Header + "Плавный пуск;Подогрев ручек;Подсветка рабочей зоны;Привод;Производительность, л/мин.;Производительность, м3/ч.;Работа от пониженного напряжения;";
            Header = Header + "Рабочее давление, бар.;Размер шлифовальной ленты, мм.;Расход топлива, кг/ч.;Расход топлива, л/ч.;Регулировка частоты вращения;Режим с ударом;";
            Header = Header + "Режимы работы;Режущая гарнитура;Реверс;Розетка для электроинструмента;Сварочный ток, А.;Температура нагрева, град.С;Тип;Тип двигателя;Тип крепления оснастки;";
            Header = Header + "Тип передвижения;Тип перекачиваемой жидкости;Тип редуктора;Тип рукоятки;Тип штанги;Толщина ведущего звена, мм.;Топливо;Уровень шума, дБ.;Фара;";
            Header = Header + "Функция TIG-сварки;Частота вращения, об/мин.;Частота ударов, уд/мин.;Шаг цепи, дюйм.;Ширина обработки, мм;Эксцентриситет;Энергия удара, Дж.";

            long num10 = Decode36("83");
            string num36 = Encode36(12919);
            string aa = RemoveLeadingNumbers("07511. Инструмент");

            string[] HeaderArray = Header.Split(';');
            Exchange1C.Exchange1C77 webClient = new Exchange1C.Exchange1C77();
            webClient.Timeout =30000000; 

            //webClient.Proxy = new WebProxy("localhost", 8888);
            //Exchange1C.FullRepairStatus crs = new Exchange1C.FullRepairStatus();
            Exchange1C.RepairStatusArray ar = new Exchange1C.RepairStatusArray();
            Exchange1C.ZakazReglamentReply pr1 = new Exchange1C.ZakazReglamentReply();
            Exchange1C.PriceList pr = new Exchange1C.PriceList();
            Console.WriteLine(DateTime.Now + ": " + "Start");
            double dk = 17174.4 + 11.4;
            try
            {
                //var f = webClient.ReportFirmaStock("xlsx", "     4S  ", new DateTime(2021,12,31));
                //if (f.ExcelFile != null)
                //{
                //    File.WriteAllBytes(@"f:\tmp\ОстаткиСтинСервис.xlsx", f.ExcelFile);
                //}
                pr1 = webClient.ReportZakazReglament(0, "xls", "01.05.2020", "11.06.2020ТА", 1, 1, 1, 1, 0, 0, 0, 0, 0, null, 1, "    PFD  ", 1,
                    "S0010,S0006,S0004,D0008,S0017,S0012,S0002,S0005,S0007,D0007",
                    null,
                    "    1MD  ,    1VD  ,    2GD  ,    2FD  ,    2DD  ,    2ED  ,    1LD  ,    24D  ",
                    "    1KD  ,    1OS  ,    27D  ,    1NS  ",
                    "    25D  ,     4S  ",
                    "    1PS  ,    2ID  ",
                    "    1TD  ,    1XD  ,    1QD  ,    21D  ,    1UD  ,    1RD  ,    1YD  ,    1WD  ",
                    null,
                    "     PS  ,     TD  ,     DS  ",
                    null);
                Exchange1C.DataEntry dataEntry1 = new Exchange1C.DataEntry { Имя = "НашаОрганизация", Значение = "Какое-то очень длинное описание с цифрами и буквами, которое не умещается на одной строке" };
                Exchange1C.DataEntry dataEntry2 = new Exchange1C.DataEntry { Имя = "Поклажедатель", Значение = "ООО Рога и копыта цифрами и буквами, которое не умещается на одной строке" };
                Exchange1C.DataEntry dataEntry3 = new Exchange1C.DataEntry { Имя = "ДатаДокумента", Значение = "19.04.2020" };
                Exchange1C.DataEntry[] dataEntries = new Exchange1C.DataEntry[] { dataEntry1, dataEntry2, dataEntry3 };

                Exchange1C.DataEntry[] dataEntries2 = new Exchange1C.DataEntry[] {
                new Exchange1C.DataEntry { Имя = "ном", Значение = "1" },
                new Exchange1C.DataEntry { Имя = "ПечТовар", Значение = "Шлифлист 93x230мм Р40 Topex 5шт" },
                new Exchange1C.DataEntry { Имя = "ПечАртикул", Значение = "63H504" }
                };
                Exchange1C.DataEntry[] dataEntries3 = new Exchange1C.DataEntry[] {
                new Exchange1C.DataEntry { Имя = "ном", Значение = "2" },
                new Exchange1C.DataEntry { Имя = "ПечТовар", Значение = "Шлифлист другой с какими-то фильдиперстывыми штучками на боку, выкрашенными в рыжий цвет" },
                new Exchange1C.DataEntry { Имя = "ПечАртикул", Значение = "01226" }
                };

                Exchange1C.nDataEntry[] nDatas = new Exchange1C.nDataEntry[] { new Exchange1C.nDataEntry { мнИмя="", мнЗначения=new Exchange1C.DataEntry[][] { dataEntries } },
                new Exchange1C.nDataEntry { мнИмя="мнТаблица", мнЗначения=new Exchange1C.DataEntry[][] { dataEntries2, dataEntries3}} };
                pr = webClient.CreateExcelByTemplate("АктСверкиСОХ.xlsx", nDatas);

                if (pr.ExcelFile != null)
                {
                    File.WriteAllBytes(@"f:\tmp\24.xls", pr.ExcelFile);
                }

                pr = webClient.GetPriceList("xls", "     DS  ", null, 0, null, 0, 0, 1, 1, 1, 1, 1, 1, 0, null, 12, "", 1, 0);
                //pr = webClient.GetFullTovarForWeb(false);
                //ar = webClient.ConsolidatedServiceReport(new DateTime(2019, 01, 14));
                //DirectoryInfo DirOrder = new DirectoryInfo(@"F:\tmp\orders");
                //var files = DirOrder.EnumerateFiles("*.csv", SearchOption.AllDirectories);
                //foreach (FileInfo file in files)
                //{
                //    string dd = file.Name;
                //    //Log("Файл " + file.Name);
                //    webClient.CreateDocuments(File.ReadAllBytes(file.FullName));
                //    //file.Delete();
                //}
                //Console.Write(webClient.CreateDocuments(File.ReadAllBytes(@"F:\tmp\order1747.csv")));
                //Console.WriteLine(webClient.GetRepairStatus("DY01662", 2016));
                //crs = webClient.GetFullRepairStatus("DY00738", 2016);
                //ar = webClient.GetAllRepairsByUserID("F0002068");
                //pr = webClient.GetFullTovarProdagiForWeb();
                string PpBrak = "     GS  ,     OS  ,    1IS  ,    1LD  ,    24D  ";
                string Brak = "     QS  ,     PS  ,     TD  ,     SS  ,     RS  ,     DS  ";
                string NomSelect = "   OQJF  "; // "   96US  ";//"   7E1S  ,  101QD  ,   VNHD  ,  156HD  ";//"   7E1S  ";//"   VNHD  ";// 
                //pr1 = webClient.ReportZakazReglament(1, "xls", "01.08.2018", "14.08.2019ТА", 1, 1, 1, 1, 0, 0, 0, 0, 0, NomSelect, 1, null, 1, "S0010,S0017,S0004,D0009,D0008", PpBrak, Brak, "    1HS  ,    1KD  ,    1AD  ,    1CD  ,    1BD  ,    1DD  ,    1FD  ,    1ED  ,    1OS  ,    1NS  ", "     4S  ", "    1WD  ,    1TD  ,    1XD  ,    1QD  ,    21D  ,    1UD  ,    1RD  ,    25D  ,    1YD  ", "     CS  ", "     NS  ", "    1PS  ","");
                File.WriteAllBytes(@"f:\tmp\test.xls", pr.ExcelFile);
            }
            finally
            {
                webClient.Dispose();
            }
            Console.WriteLine(DateTime.Now + ": " + "Ready");
            Console.ReadKey();
        }

        private static byte[] GluingFiles(FileInfo file)
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

        private static Int64 Decode36(string input)
        {
            string CharList = "0123456789abcdefghijklmnopqrstuvwxyz";
            var reversed = input.ToLower().Reverse();
            long result = 0;
            int pos = 0;
            foreach (char c in reversed)
            {
                result += CharList.IndexOf(c) * (long)Math.Pow(36, pos);
                pos++;
            }
            return result;
        }

        private static String Encode36(long input)
        {
            if (input < 0)
                return "";
            else
            {
                string CharList = "0123456789abcdefghijklmnopqrstuvwxyz";
                char[] clistarr = CharList.ToUpper().ToCharArray();
                var result = new Stack<char>();
                while (input != 0)
                {
                    result.Push(clistarr[input % 36]);
                    input /= 36;
                }
                return new string(result.ToArray());
            }
        }
        static string RemoveLeadingNumbers(string AString)
        {
            if (string.IsNullOrEmpty(AString))
                return AString;
            AString = AString.Trim();
            int n = 0;
            while (Char.IsNumber(AString[n]))
                n++;
            if (n > 0)
            {
                if ((AString[n] == '.') | (AString[n] == '-'))
                    n++;
                AString = AString.Substring(n);
            }
            return AString.Trim();
        }

    }
}
