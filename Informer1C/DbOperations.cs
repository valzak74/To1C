using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Core.Objects;
using System.Collections.ObjectModel;

namespace Informer1C
{
    public static class DbOperations
    {
        private static Models.DbModel _dbContext = null;
        private static string _lastDocZ;
        private const string RobotId = "    20D  ";
        private const string ПустоеЗначение = "     0   ";

        private static bool ConnectionAccepted()
        {
            using (Models.SysModel sysContext = new Models.SysModel())
            {
                System.Data.SqlClient.SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder(System.Configuration.ConfigurationManager.ConnectionStrings["DbModel"].ConnectionString);
                string Db = builder.InitialCatalog;
                
                sysContext.Database.Connection.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["SysModel"].ConnectionString;
                sysContext.Database.Connection.Open();

                return (int)sysContext.Database.SqlQuery<int>("select count(*) from sys.sysprocesses where db_name(dbid) = '" + Db + "'").FirstOrDefault() > 0;
            }
        }

        private static long Decode36(string input)
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

        private static string Encode36(long input)
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

        private static DateTime TryGetDate(string value)
        {
            DateTime d;
            long ticks = 0;
            if (value.Length > 8)
            {
                ticks = Decode36(value.Substring(8, 6)) * 1000;
                value = value.Substring(0, 8);
            }
            if (!DateTime.TryParseExact(value, "yyyyMMdd", new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, out d))
                d = DateTime.MinValue;
            return d.AddTicks(ticks);
        }

        private static string GenerateIdDoc(Models.DbModel dbContext)
        {
            string num36_pf = "";
            long num10 = 0;
            string num36 = dbContext.C_1SJOURN.Max(x => x.IDDOC);
            num36_pf = num36.Substring(6);
            num36 = num36.Substring(0, 6).Trim();
            if (!string.IsNullOrEmpty(num36))
                num10 = Decode36(num36);
            num10++;
            return (Encode36(num10) + num36_pf).PadLeft(9);
        }

        private static string GenerateDocNo(Models.DbModel dbContext, int IdDocDef, string DocPref)
        {
            string yyyy = DateTime.Today.ToString("yyyy");
            string DocNo = dbContext.C_1SJOURN.Where(x => x.IDDOCDEF == IdDocDef & x.DOCNO.Substring(0, 2) == DocPref & x.DATE_TIME_IDDOC.Substring(0, 4) == yyyy).Max(d => d.DOCNO);
            if (string.IsNullOrEmpty(DocNo))
                return "1";
            else
            {
                return (Convert.ToInt32(DocNo.Substring(2)) + 1).ToString();
            }
        }

        private static FirmaDetails GetFirma(Models.DbModel dbContext)
        {
            return
                (from p in dbContext.SC30.Where(x => x.ID == RobotId)
                join f in dbContext.SC4014 on p.SP4010 equals f.ID
                join u in dbContext.SC131 on f.SP4011 equals u.ID
                select new FirmaDetails
                {
                    Id = f.ID,
                    Name = f.DESCR,
                    UrLitso = f.SP4011,
                    Prefix = u.SP145.Trim(),
                    BankSchet = f.SP4133
                }).FirstOrDefault();
        }

        private static bool CanCreateDoc(Models.DbModel dbContext, int IdDocDef, string DocNo)
        {
            string yyyy = DateTime.Today.ToString("yyyy");
            return !dbContext.C_1SJOURN.Any(x => x.IDDOCDEF == IdDocDef & x.DOCNO == DocNo & x.DATE_TIME_IDDOC.Substring(0, 4) == yyyy);
        }

        public static Models.DbModel GetDbContext()
        {
            if ((_dbContext == null) && ConnectionAccepted())
            {
                _dbContext = new Models.DbModel();
                _dbContext.Database.Connection.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DbModel"].ConnectionString;
                _dbContext.Database.Connection.Open();
            }

            return _dbContext;
        }

        public static void FreeDbContext()
        {
            if (_dbContext != null)
            {
                _dbContext.Database.Connection.Close();
                _dbContext = null;
            }
        }

        public static bool DbContextIsNull()
        {
            return _dbContext == null;
        }

        // Преобразование IQueryable в BindingList
        public static BindingList<T> ToBindingList<T>
            (this IQueryable<T> source) where T : class
        {
            return (new ObservableCollection<T>(source)).ToBindingList();
        }

        // Обновление всех изменённых объектов в коллекции
        public static void Refresh(this Models.DbModel dbContext, RefreshMode mode,
           IEnumerable collection)
        {
            var objectContext = ((IObjectContextAdapter)dbContext).ObjectContext;
            objectContext.Refresh(mode, collection);
        }

        // Обновление объекта
        public static void Refresh(this Models.DbModel dbContext, RefreshMode mode,
             object entity)
        {
            var objectContext = ((IObjectContextAdapter)dbContext).ObjectContext;
            objectContext.Refresh(mode, entity);
        }

        public static IQueryable<Заявка> GetData(List<string> Производители)
        {
            Models.DbModel dbContext = GetDbContext();
            if (dbContext == null)
                return null;
            DateTime Period = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            
            var tbl =
                from rg in dbContext.RG4667
                join spNom in dbContext.SC84 on rg.SP4663 equals spNom.ID
                where rg.PERIOD == Period & rg.SP13435 == 1 & Производители.Contains(spNom.SP8842) & rg.SP4666 != 0
                group rg by rg.SP4664 into regZZ
                join docZ in dbContext.DH2457 on regZZ.FirstOrDefault().SP4664 equals docZ.IDDOC
                join j in dbContext.C_1SJOURN on docZ.IDDOC equals j.IDDOC
                join spPol in dbContext.SC30 on j.SP74 equals spPol.ID
                join spClient in dbContext.SC172 on docZ.SP2434 equals spClient.ID
                select new Заявка
                {
                    IDDOC = docZ.IDDOC,
                    НомерДок = j.DOCNO,
                    ДатаДок = j.DATE_TIME_IDDOC.Substring(0, 14),
                    НомерДокПолный = j.DOCNO + " от " + j.DATE_TIME_IDDOC.Substring(6, 2) + "." + j.DATE_TIME_IDDOC.Substring(4, 2) + "." + j.DATE_TIME_IDDOC.Substring(0, 4),
                    Автор = spPol.DESCR,
                    КонтрагентID = docZ.SP2434,
                    Контрагент = spClient.DESCR,
                    СтрокиДокумента = (
                        from docZDetail in dbContext.DT2457.Where(x => x.IDDOC == docZ.IDDOC).GroupBy(x => x.SP2446)
                        join spNom in dbContext.SC84 on docZDetail.FirstOrDefault().SP2446 equals spNom.ID
                        join spBrand in dbContext.SC8840 on spNom.SP8842 equals spBrand.ID
                        where Производители.Contains(spBrand.ID)
                        select new ЗаявкаСтрока
                        {
                            НоменклатураID = spNom.ID,
                            Номенклатура = spNom.DESCR,
                            Артикул = spNom.SP85,
                            ПроизводительID = spBrand.ID,
                            Производитель = spBrand.DESCR,
                            Количество = docZDetail.Sum(x => (decimal?)x.SP2447),
                            ЦенаЗаявленная = docZDetail.FirstOrDefault().SP2450,
                            ЦенаСогласованная = docZDetail.FirstOrDefault().SP2450,
                            СрокПоставки = 1,
                            Действие = (int)Вариант.НеОбработана
                        }).ToList()
                };

            _lastDocZ = tbl.Select(x => x.IDDOC).Max();

            return tbl;
        }

        public static List<НоваяЗаявка> GetFreshData(List<string> Производители)
        {
            Models.DbModel dbContext = GetDbContext();
            if (dbContext == null)
                return null;
            DateTime Period = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

            var tbl =
                from rg in dbContext.RG4667
                join spNom in dbContext.SC84 on rg.SP4663 equals spNom.ID
                where rg.PERIOD == Period & rg.SP13435 == 1 & Производители.Contains(spNom.SP8842) & string.Compare(rg.SP4664, _lastDocZ) == 1 & rg.SP4666 != 0
                group rg by rg.SP4664 into regZZ
                join docZ in dbContext.DH2457 on regZZ.FirstOrDefault().SP4664 equals docZ.IDDOC
                join j in dbContext.C_1SJOURN on docZ.IDDOC equals j.IDDOC
                join spClient in dbContext.SC172 on docZ.SP2434 equals spClient.ID
                select new НоваяЗаявка
                {
                    IDDOC = docZ.IDDOC,
                    Контрагент = spClient.DESCR,
                    НомерДокПолный = j.DOCNO + " от " + j.DATE_TIME_IDDOC.Substring(6, 2) + "." + j.DATE_TIME_IDDOC.Substring(4, 2) + "." + j.DATE_TIME_IDDOC.Substring(0, 4)
                };

            var newList = tbl.ToList();

            if (newList.Count() > 0)
                _lastDocZ = newList.Select(x => x.IDDOC).Max();

            return newList;
        }

        public static bool DataChanged(BindingList<Заявка> Data)
        {
            return Data.Where(x => x.СтрокиДокумента.Any(y => y.Действие != Вариант.НеОбработана)).Count() > 0;
        }

        public static void CreateDocument(BindingList<Заявка> Data)
        {
            var ChangedData = Data.Where(x => x.СтрокиДокумента.Any(y => y.Действие != Вариант.НеОбработана)).ToList();
            
            if (ChangedData.Count > 0)
            {
                try
                {

                    Models.DbModel dbContext = GetDbContext();
                    if (dbContext != null)
                    {
                        int iddocdef = 13433; //СогласованиеЗаявок
                        //1sjourn
                        dbContext.Refresh(RefreshMode.StoreWins, dbContext.ChangeTracker.Entries().Select(c => c.Entity).ToList());
                        string num36 = GenerateIdDoc(dbContext);
                        FirmaDetails firma = GetFirma(dbContext);
                        string DocNo = "E" + firma.Prefix + GenerateDocNo(dbContext, iddocdef, "E" + firma.Prefix).PadLeft(8, '0'); // "EW00000001"; 
                        if (CanCreateDoc(dbContext, iddocdef, DocNo))
                        {
                            var RowJourn = new Models.C_1SJOURN();
                            RowJourn.ROW_ID = dbContext.C_1SJOURN.Max(x => x.ROW_ID) + 1;
                            RowJourn.IDJOURNAL = 0;
                            RowJourn.IDDOC = num36;
                            RowJourn.IDDOCDEF = iddocdef;
                            RowJourn.APPCODE = 1;
                            RowJourn.DATE_TIME_IDDOC = DateTime.Today.ToString("yyyyMMdd") + Encode36(180000000) + num36;
                            RowJourn.DNPREFIX = iddocdef.ToString().PadLeft(10) + DateTime.Today.ToString("yyyy").PadRight(8);
                            RowJourn.DOCNO = DocNo;
                            RowJourn.CLOSED = 0;
                            RowJourn.ISMARK = false;
                            RowJourn.ACTCNT = 0;
                            RowJourn.VERSTAMP = 0;

                            RowJourn.SP74 = RobotId;
                            RowJourn.SP798 = ПустоеЗначение;
                            RowJourn.SP4056 = firma.Id;
                            RowJourn.SP5365 = firma.UrLitso;
                            RowJourn.SP8662 = "D";
                            RowJourn.SP8663 = "D;Экран";
                            RowJourn.SP8664 = "D";
                            RowJourn.SP8665 = "D;СогласованиеЗаявок";
                            RowJourn.SP8666 = "D;" + (firma.Name.Length > 28 ? firma.Name.Substring(0, 28) : firma.Name);
                            RowJourn.SP8720 = "D";
                            RowJourn.SP8723 = "D";

                            RowJourn.DS1946 = 0;
                            RowJourn.DS4757 = 0;
                            RowJourn.DS5722 = 0;

                            dbContext.C_1SJOURN.Add(RowJourn);

                            //dh13433
                            var DocСогласованиеЗаявок = new Models.DH13433();
                            DocСогласованиеЗаявок.IDDOC = num36;
                            DocСогласованиеЗаявок.SP13424 = ПустоеЗначение;
                            DocСогласованиеЗаявок.SP660 = "Сформирован автоматически";
                            dbContext.DH13433.Add(DocСогласованиеЗаявок);

                            short lineno = 0;
                            foreach (Заявка doc in ChangedData)
                            {
                                foreach (ЗаявкаСтрока str in doc.СтрокиДокумента)
                                {
                                    //dt13433
                                    lineno++;
                                    var DocМнСогласованиеЗаявок = new Models.DT13433();
                                    DocМнСогласованиеЗаявок.IDDOC = num36;
                                    DocМнСогласованиеЗаявок.LINENO_ = lineno;
                                    DocМнСогласованиеЗаявок.SP13425 = doc.IDDOC;
                                    DocМнСогласованиеЗаявок.SP13426 = str.НоменклатураID;
                                    DocМнСогласованиеЗаявок.SP13427 = (decimal)str.Количество;
                                    DocМнСогласованиеЗаявок.SP13428 = (decimal)str.ЦенаЗаявленная;
                                    DocМнСогласованиеЗаявок.SP13429 = (decimal)str.ЦенаСогласованная;
                                    DocМнСогласованиеЗаявок.SP13430 = (decimal)str.СрокПоставки;
                                    DocМнСогласованиеЗаявок.SP13431 = (decimal)str.Действие;

                                    dbContext.DT13433.Add(DocМнСогласованиеЗаявок);

                                    //_1scrdoc
                                    var ПодчДокументы = new Models.C_1SCRDOC();
                                    ПодчДокументы.ROW_ID = dbContext.C_1SCRDOC.Max(x => x.ROW_ID) + 1;
                                    ПодчДокументы.MDID = 0;
                                    ПодчДокументы.PARENTVAL = "O1 1W9" + doc.IDDOC;
                                    ПодчДокументы.CHILD_DATE_TIME_IDDOC = DateTime.Today.ToString("yyyyMMdd") + Encode36(180000000) + num36;
                                    ПодчДокументы.CHILDID = num36;
                                    ПодчДокументы.FLAGS = 1;

                                    dbContext.C_1SCRDOC.Add(ПодчДокументы);

                                }
                            }

                            dbContext.SaveChanges();
                        }
                    }
                }
                catch (DbEntityValidationException e)
                {
                    foreach (var eve in e.EntityValidationErrors)
                    {
                        Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                            eve.Entry.Entity.GetType().Name, eve.Entry.State);
                        foreach (var ve in eve.ValidationErrors)
                        {
                            Console.WriteLine("- Property: \"{0}\", Value: \"{1}\", Error: \"{2}\"",
                                ve.PropertyName,
                                eve.Entry.CurrentValues.GetValue<object>(ve.PropertyName),
                                ve.ErrorMessage);
                        }
                    }
                    throw;
                }
            }
        }
    }
}
