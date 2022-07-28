using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Linq;
using System.IO;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;

namespace To1C
{
    /// <summary>
    /// Summary description for Doc77Loader
    /// </summary>
    [WebService(Namespace = "http://stin-base.ru/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class Exchange1C77 : System.Web.Services.WebService
    {
        [WebMethod(Description = "Получение короткого значения статуса ремонта")]
        public string GetRepairStatus(string KvNumber, int KvDate)
        {
            string FunctionResult = "Не обнаружен в БД";
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = @"SELECT sp9963 FROM rg9972 (NOLOCK) WHERE (sp9969 = @KvitNumber) and (sp10084 = @KvitDate) and (period = @Period) and (sp9970 > 0)";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Parameters.Add("@KvitNumber", SqlDbType.Char);
                command.Parameters.Add("@KvitDate", SqlDbType.Int);
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                command.Parameters["@KvitNumber"].Value = KvNumber;
                command.Parameters["@KvitDate"].Value = KvDate;
                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                command.Parameters["@Period"].Value = startDate;

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        FunctionResult = GetStatus(reader.GetString(0).Trim());
                    }

                }
                catch (Exception e)
                {
                    FunctionResult = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }

            return FunctionResult;
        }

        [WebMethod(Description = "Получение развернутого значения статуса ремонта")]
        public FullRepairStatus GetFullRepairStatus(string KvNumber, int KvDate)
        {
            FullRepairStatus FullRS = new FullRepairStatus();
            FullRS.ReplyState = new ReplyState();
            FullRS.ReplyState.State = "Error";
            FullRS.ReplyState.Description = "Не обнаружен в БД";
            FullRS.CompleteRepairStatus = new CompleteRepairStatus();
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = @"SELECT pm.sp9963, pm.sp9969, pm.sp10084, pm.sp9958, pm.sp9961, pm.sp9967, k.descr, n.descr,
                                        docRez.SP11032, docZaver.SP10455
                                        FROM rg9972 as pm (NOLOCK)
                                        INNER JOIN sc172 as k (NOLOCK) on (pm.sp9964 = k.id)
                                        INNER JOIN sc84 as n (NOLOCK) on (pm.sp9960 = n.id)
                                        left join DH11037 as docRez (NOLOCK) on ((pm.sp9969 = docRez.SP11006) and (pm.sp10084 = docRez.SP11007))
                                        left join RG10476 as rZaver (NOLOCK) on ((pm.sp9969 = rZaver.SP10472) and (pm.sp10084 = rZaver.SP10473) and (rZaver.SP10475 > 0))
                                        left join DH10457 as docZaver (NOLOCK) on (rZaver.SP10474 = docZaver.IDDOC)
                                        WHERE (pm.sp9969 = @KvitNumber) and (pm.sp10084 = @KvitDate) and (pm.period = @Period) and (pm.sp9970 > 0)";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Parameters.Add("@KvitNumber", SqlDbType.Char);
                command.Parameters.Add("@KvitDate", SqlDbType.Int);
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                command.Parameters["@KvitNumber"].Value = KvNumber;
                command.Parameters["@KvitDate"].Value = KvDate;
                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                command.Parameters["@Period"].Value = startDate;

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        FullRS.CompleteRepairStatus.Status = GetStatus(reader.GetString(0).Trim());
                        FullRS.CompleteRepairStatus.KvitNumber = reader.GetString(1).Trim();
                        FullRS.CompleteRepairStatus.KvitDate = Convert.ToInt16(reader.GetValue(2));
                        FullRS.CompleteRepairStatus.RepairType = GetRepairType(Convert.ToInt16(reader.GetValue(3)));
                        FullRS.CompleteRepairStatus.ZavNumber = reader.GetString(4).Trim();
                        FullRS.CompleteRepairStatus.DatePriema = reader.GetDateTime(5);
                        FullRS.CompleteRepairStatus.Kontragent = reader.GetString(6).Trim();
                        FullRS.CompleteRepairStatus.Izdelie = reader.GetString(7).Trim();

                        if (!reader.IsDBNull(8) && ((FullRS.CompleteRepairStatus.RepairType == "Платный") || ((FullRS.CompleteRepairStatus.RepairType == "Экспертиза") && (FullRS.CompleteRepairStatus.Status == "Экспертиза завершена"))))
                        {
                            FullRS.CompleteRepairStatus.PredvarSumma = reader.GetDecimal(8);
                        }
                        if (!reader.IsDBNull(9) && (FullRS.CompleteRepairStatus.RepairType == "Платный"))
                        {
                            FullRS.CompleteRepairStatus.Summa = reader.GetDecimal(9);
                        }

                        FullRS.ReplyState.State = "Success";
                        FullRS.ReplyState.Description = "Информация успешно извлечена из БД";
                    }

                }
                catch (Exception e)
                {
                    FullRS.ReplyState.State = "Error";
                    FullRS.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }

            return FullRS;
        }

        [WebMethod(Description = "Получение развернутого статуса всех ремонтов клиента")]
        public RepairStatusArray GetAllRepairsByUserID(string UserID)
        {
            RepairStatusArray rsArray = new RepairStatusArray();
            rsArray.ReplyState = new ReplyState();
            rsArray.ReplyState.State = "Error";
            rsArray.ReplyState.Description = "Не обнаружен в БД";
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = @"SELECT pm.sp9963, pm.sp9969, pm.sp10084, pm.sp9958, pm.sp9961, pm.sp9967, k.descr, n.descr,
                                        docRez.SP11032, docZaver.SP10455
                                        FROM rg9972 as pm (NOLOCK)
                                        INNER JOIN sc172 as k (NOLOCK) on (pm.sp9964 = k.id)
                                        INNER JOIN sc84 as n (NOLOCK) on (pm.sp9960 = n.id)
                                        left join DH11037 as docRez (NOLOCK) on ((pm.sp9969 = docRez.SP11006) and (pm.sp10084 = docRez.SP11007))
                                        left join RG10476 as rZaver (NOLOCK) on ((pm.sp9969 = rZaver.SP10472) and (pm.sp10084 = rZaver.SP10473) and (rZaver.SP10475 > 0))
                                        left join DH10457 as docZaver (NOLOCK) on (rZaver.SP10474 = docZaver.IDDOC)
                                        WHERE (k.code = @Code) and (pm.period = @Period) and (pm.sp9970 > 0)
                                        ORDER BY pm.sp9967";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Parameters.Add("@Code", SqlDbType.Char);
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                command.Parameters["@Code"].Value = UserID;
                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                command.Parameters["@Period"].Value = startDate;

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                List<CompleteRepairStatus> rs = new List<CompleteRepairStatus>();
                try
                {
                    while (reader.Read())
                    {
                        CompleteRepairStatus crs = new CompleteRepairStatus();
                        crs.Status = GetStatus(reader.GetString(0).Trim());
                        crs.KvitNumber = reader.GetString(1).Trim();
                        crs.KvitDate = Convert.ToInt16(reader.GetValue(2));
                        crs.RepairType = GetRepairType(Convert.ToInt16(reader.GetValue(3)));
                        crs.ZavNumber = reader.GetString(4).Trim();
                        crs.DatePriema = reader.GetDateTime(5);
                        crs.Kontragent = reader.GetString(6).Trim();
                        crs.Izdelie = reader.GetString(7).Trim();

                        if (!reader.IsDBNull(8) && ((crs.RepairType == "Платный") || ((crs.RepairType == "Экспертиза") && (crs.Status == "Экспертиза завершена"))))
                        {
                            crs.PredvarSumma = reader.GetDecimal(8);
                        }
                        if (!reader.IsDBNull(9) && (crs.RepairType == "Платный"))
                        {
                            crs.Summa = reader.GetDecimal(9);
                        }
                        rs.Add(crs);
                    }

                    rsArray.RepairStatuses = rs.ToArray();

                    rsArray.ReplyState.State = "Success";
                    rsArray.ReplyState.Description = "Информация успешно извлечена из БД";
                }
                catch (Exception e)
                {
                    rsArray.ReplyState.State = "Error";
                    rsArray.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }
            return rsArray;
        }

        private string connStr = ConfigurationManager.ConnectionStrings["StinDB"].ConnectionString;
        private string connDocsStr = ConfigurationManager.ConnectionStrings["DocsDB"].ConnectionString;
        private string connWebStr = ConfigurationManager.ConnectionStrings["WebDB"].ConnectionString;

        private string GetStatus(string StatusDB)
        {
            string FunctionResult = "";
            switch (StatusDB)
            {
                case "7OH":
                    FunctionResult = "Принят в ремонт";
                    break;
                case "7OI":
                    FunctionResult = "Готов к выдаче";
                    break;
                case "8IM":
                    FunctionResult = "На диагностике";
                    break;
                case "8KG":
                    FunctionResult = "Экспертиза завершена";
                    break;
                case "8KH":
                    FunctionResult = "Отказ от платного ремонта";
                    break;
                case "8IN":
                    FunctionResult = "Ожидание запасных частей";
                    break;
                case "8IO":
                    FunctionResult = "В ремонте";
                    break;
                default:
                    FunctionResult = "НЕ ОБНАРУЖЕН";
                    break;
            }
            return FunctionResult;
        }

        private string GetRepairType(int RepairTypeDB)
        {
            string FunctionResult = "";
            switch (RepairTypeDB)
            {
                case 0:
                    FunctionResult = "Платный";
                    break;
                case 1:
                    FunctionResult = "Гарантийный";
                    break;
                case 2:
                    FunctionResult = "Предпродажный";
                    break;
                case 3:
                    FunctionResult = "За свой счет";
                    break;
                case 4:
                    FunctionResult = "Экспертиза";
                    break;
                default:
                    FunctionResult = "НЕ ОБНАРУЖЕН";
                    break;
            }
            return FunctionResult;
        }

        private void SetTree(List<Price> DestList, List<Price> DataList, string CurrentID)
        {
            foreach (Price node in DataList.Where(child => child.ParentID == CurrentID).OrderBy(child => child.IsFolder).ToList())
            {
                DestList.Add(node);
                SetTree(DestList, DataList, node.ID);
            }
        }

        [WebMethod(Description = "Устанавливает флаг ГОТОВ в док-те Набор")]
        public string SetClosedFlagVNabore(string iddoc, int closed, string worker, int f0, int f1, int f2)
        {
            string functionResult = "";
            if (!string.IsNullOrEmpty(iddoc))
            {
                using (SqlConnection connection = new SqlConnection(connDocsStr))
                {
                    connection.Open();

                    string docdate = "";
                    string curperiod = "";
                    int actno = 0;
                    int rf12566 = 0;
                    string queryString = "select ACTCNT, DATE_TIME_IDDOC, RF12566 from _1sjourn (NOLOCK) where iddoc=@iddoc;";
                    SqlCommand commandsel = new SqlCommand(queryString, connection);
                    commandsel.Parameters.AddWithValue("@iddoc", iddoc);
                    SqlDataReader reader = commandsel.ExecuteReader();

                    try
                    {
                        while (reader.Read())
                        {
                            actno = reader.GetInt32(0);
                            docdate = reader.GetString(1).Trim().Substring(0, 8);
                            curperiod = docdate.Substring(0, 6) + "01";
                            rf12566 = Convert.ToInt32(reader.GetValue(2));
                        }
                    }
                    catch (Exception e)
                    {
                        functionResult = e.Message;
                    }
                    finally
                    {
                        reader.Close();
                        commandsel.Dispose();
                    }
                    if ((!string.IsNullOrEmpty(docdate)) && (!string.IsNullOrEmpty(curperiod)))
                    {
                        string query = "update DH11948 set SP11938 = @closed, SP12559 = @worker where iddoc = @iddoc;";
                        if (string.IsNullOrEmpty(worker))
                            worker = "     0   ";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@iddoc", iddoc);
                            command.Parameters.AddWithValue("@closed", closed);
                            command.Parameters.AddWithValue("@worker", worker);
                            command.ExecuteNonQuery();
                        }
                        if (rf12566 == 1)
                        {
                            actno = actno - 1;
                            using (SqlCommand commandReg = new SqlCommand("_1sp_RA12566_ClearRecalcDocAct", connection)
                            {
                                CommandType = CommandType.StoredProcedure
                            })
                            {
                                commandReg.Parameters.AddWithValue("@IdDoc", iddoc);
                                commandReg.Parameters.AddWithValue("@DocDate", docdate);
                                commandReg.Parameters.AddWithValue("@CurPeriod", curperiod);
                                commandReg.Parameters.AddWithValue("@RepeatToTM", 0);
                                commandReg.Parameters.AddWithValue("@SaveTurnsWithMonth", 0);
                                commandReg.Parameters.AddWithValue("@Direct", 1);
                                commandReg.ExecuteNonQuery();
                            }
                            query = @"delete from ra12566 where iddoc = @iddoc
                                        update _1sjourn set actcnt = @actcnt, rf12566 = 0 where iddoc=@iddoc";
                            using (SqlCommand commandUpdate = new SqlCommand(query, connection))
                            {
                                commandUpdate.Parameters.AddWithValue("@iddoc", iddoc);
                                commandUpdate.Parameters.AddWithValue("@actcnt", actno);
                                commandUpdate.ExecuteNonQuery();
                            }
                        }
                        if ((worker != "     0   ") && (closed == 1))
                        {
                            actno = actno + 1;
                            using (SqlCommand commandReg = new SqlCommand("_1sp_RA12566_WriteDocAct", connection)
                            {
                                CommandType = CommandType.StoredProcedure
                            })
                            {
                                commandReg.Parameters.AddWithValue("@IdDoc", iddoc);
                                commandReg.Parameters.AddWithValue("@LineNo", 0);
                                commandReg.Parameters.AddWithValue("@ActNo", actno);
                                commandReg.Parameters.AddWithValue("@DebetCredit", 0);
                                commandReg.Parameters.AddWithValue("@p0", worker);
                                commandReg.Parameters.AddWithValue("@f0", f0);
                                commandReg.Parameters.AddWithValue("@f1", f1);
                                commandReg.Parameters.AddWithValue("@f2", f2);
                                commandReg.Parameters.AddWithValue("@DocDate", docdate);
                                commandReg.Parameters.AddWithValue("@CurPeriod", curperiod);
                                commandReg.Parameters.AddWithValue("@RepeatToTM", 0);
                                commandReg.Parameters.AddWithValue("@SaveTurnsWithMonth", 0);
                                commandReg.ExecuteNonQuery();
                            }
                            query = "update _1sjourn set actcnt = @actcnt, rf12566 = 1 where iddoc=@iddoc";
                            using (SqlCommand commandjourn = new SqlCommand(query, connection))
                            {
                                commandjourn.Parameters.AddWithValue("@iddoc", iddoc);
                                commandjourn.Parameters.AddWithValue("@actcnt", actno);
                                commandjourn.ExecuteNonQuery();
                            }
                        }
                        if (string.IsNullOrEmpty(functionResult))
                            functionResult = "SUCCESS";
                    }
                    else
                        if (string.IsNullOrEmpty(functionResult))
                        functionResult = "Не найдены данные в _1sjourn";

                    return functionResult;
                }
            }
            else
                return "IDDOC is empty";
        }

        public void LogMessageToFile(string msg)
        {
            System.IO.StreamWriter sw = System.IO.File.AppendText(@"f:\tmp\LogTo1C.txt");
            try
            {
                string logLine = System.String.Format(
                    "{0:G}: {1}.", System.DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        [WebMethod(Description = "Получение прайс-листа в формате Excel")]
        public PriceList GetPriceList(string ext, string SkladListStr = null, string NomkListStr = null, int NomInSelect = 0, string BrendListStr = null, int BrendInSelect = 0,
                                        int harakter = 0, int opt = 0, int osob = 0, int spec = 0, int rozn = 0, int zakaz = 0, int ostPrice = 0, int onlyOst = 1, string CustomerListStr = null, int PeriodCode = 12, string Address = "", int Top = 0, int IsMasterskaya = 0)
        {
            PriceList PrList = new PriceList();
            PrList.ReplyState = new ReplyState();
            PrList.ReplyState.State = "Error";
            PrList.ReplyState.Description = "Нет информации в БД";

            List<string> listUsedNomk = new List<string>();
            bool NeedToCalcCost = false;
            if (!string.IsNullOrEmpty(CustomerListStr))
            {
                NeedToCalcCost = true;
                NomenklaturaProdanaReply npr = NomenklaturaProdana(CustomerListStr);
                if (npr.ReplyState.State == "Success")
                {
                    foreach (CustTovar t in npr.Tovars)
                    {
                        listUsedNomk.Add(t.Id);
                    }
                }
            }

            List<string> listSklad = new List<string>(SkladListStr.Split(','));
            if ((listSklad.Count == 0) || ((listSklad.Count == 1) && (listSklad[0] == "")))
            {
                PrList.ReplyState.Description = "Необходимо указать склад для вывода остатков";
                return PrList;
            }
            listSklad.Sort();
            List<string> listNomk;
            List<string> listBrend;
            if (String.IsNullOrEmpty(NomkListStr))
                listNomk = new List<string>();
            else
                listNomk = new List<string>(NomkListStr.Split(','));
            if ((listNomk.Count == 1) && (listNomk[0].Trim() == ""))
                listNomk.Clear();
            if (String.IsNullOrEmpty(BrendListStr))
                listBrend = new List<string>();
            else
                listBrend = new List<string>(BrendListStr.Split(','));
            if ((listBrend.Count == 1) && (listBrend[0].Trim() == ""))
                listBrend.Clear();

            string where = "";

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = "";
                if (NeedToCalcCost)
                {
                    //порог себестоимости
                    queryString += @"select value from _1SCONST(NOLOCK) where id = 12572;";
                    //условия контрагента
                    queryString += @"select SC172.id, SC172.code, SC172.descr, SC172.SP667, ISNULL(sc204.SP1948, '     4   '), ISNULL(sc204.SP1920,'     0   '), SC172.SP9631, SC172.SP12916, ISNULL(SC9633.SP11782,'     0   '), ISNULL(SC426.SP429, 0), IsNull(sc9678.sp9696, 0) from SC172 (NOLOCK)
                                                  left join SC204 (NOLOCK) on sc204.id = sc172.sp667
                                                  left join SC426 (NOLOCK) on SC426.id = sc204.SP1920
                                                  left join SC9633 (NOLOCK) on SC9633.id = sc172.SP9631
                                                  left join SC9678 (NOLOCK) on sc9678.id = sc204.sp9664
                                                  where SC172.code in ({3}); ";
                    //колонки скидок
                    queryString += @"select distinct SP11784,SP11785,SP11786,SP11787,SP11788,SP11789 from SC11791 (NOLOCK)
                                    where IsMark = 0;";
                    if (IsMasterskaya == 0)
                    {
                        //доп условия клиента
                        queryString += @"select SC172.id, SC9671.SP9667, SC426.SP429 from SC9671 (NOLOCK)
                                    inner join SC172 (NOLOCK) on SC172.id = SC9671.parentext
                                    inner join SC426 (NOLOCK) on SC426.id = SC9671.SP9668 
                                    where SC9671.ismark = 0 and SC172.code in ({3}) and SC9671.SP11255 = 1;";
                    }
                    else
                    {
                        //доп условия клиента (мастерская)
                        queryString += @"select SC172.id, SC12703.SP12698, SC426.SP429 from SC12703 (NOLOCK)
                                    inner join SC172 (NOLOCK) on SC172.id = SC12703.parentext
                                    inner join SC426 (NOLOCK) on SC426.id = SC12703.SP12699 
                                    where SC12703.ismark = 0 and SC172.code in ({3}) and SC12703.SP12701 = 1;";
                    }
                }
                queryString += @"select descr from sc55 (NOLOCK) where id in ({0}) order by id;";
                queryString += @" 
                                        SELECT 
                                        Gr1.ID, Gr1.Descr, 
                                        Gr2.ID, Gr2.Descr, 
                                        Gr3.ID, Gr3.Descr,
                                        Sp.ID, Sp.sp85, Sp.Descr, SpBrend.Descr, Sp.SP8848, 
                                        R.Ostatok, R.OstatokPrihod, R.OstatokPartii, R.SummaPartii, IsNull(SpBrend.id,'')";
                for (int i = 0; i < listSklad.Count; i++)
                {
                    queryString = queryString + @", ISNULL(R_Skl" + i.ToString() + ".ostatok,0)";
                }
                if ((opt == 1) || (osob == 1))
                {
                    queryString = queryString + @", ISNULL(TabConst.VALUE, 0), IsNull(sc219.sp221, 0), IsNull(SpBrend.SP11668, 0)";
                }
                if (spec == 1)
                {
                    queryString = queryString + @", ISNULL(TabConst4.VALUE, 0)";
                    if (rozn == 1)
                        queryString = queryString + @", ISNULL(TabConst5.VALUE, 0)";
                }
                if (rozn == 1)
                    queryString = queryString + @", ISNULL(TabConstRoz.VALUE, 0)";
                queryString += ", Sp.SP10397, SpEd.SP78 as koff";
                if (NeedToCalcCost)
                    queryString += ", ISNULL(TabConstZakup.VALUE, 0) as costZakup";
                if (onlyOst == 1)
                {
                    queryString = queryString + @"
                                        From
                                        (
                                            select nomk, Sum(kolvo) as Ostatok, Sum(kolvoPrih) as OstatokPrihod, Sum(kolPartii) as OstatokPartii, Sum(summaPartii) as SummaPartii from 
                                            (
                                            select sp408 as nomk, sp411 as kolvo, 0 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from rg405 (NOLOCK)
                                            where period = @Period and SP418 in (select sc55.id from sc55 (NOLOCK) where sc55.code in ('S0010','S0007','S0013','S0002','S0004','S0006','S0005','D0002','D0001'))
                                            Union all
                                            select SP8698 as nomk, SP8701 as kolvo, 0 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from RG8696 (NOLOCK)
                                            where period = @Period and SP8699 in (select sc55.id from sc55 (NOLOCK) where sc55.code in ('S0010','S0007','S0013','S0002','S0004','S0006','S0005','D0002','D0001'))
                                            Union all
                                            select SP466 as nomk, 0 as kolvo, SP4471 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from RG464 (NOLOCK)
                                            where period = @Period";
                    if (NeedToCalcCost)
                        queryString += @"
                                            Union all
                                            select SP331 as nomk, 0 as kolvo, 0 as kolvoPrih, SP342 as kolPartii, SP421 as summaPartii from RG328 (NOLOCK) 
                                            where period = @Period and SP4061 in ('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + "')";
                    queryString += @"
                                            ) x
                                            group by nomk
                                        ) as R
         
                                        inner join sc84 as Sp (NOLOCK) on (Sp.id = R.nomk) and (Sp.IsMark = 0) and (Sp.SP5066 = 0) " + (Top == 1 ? "and (Sp.SP13277 = 1) " : "") + @"and (sp.id not in (
                                            SELECT Mus.ID FROM SC84 as Mus (NOLOCK) WHERE (Mus.ISMARK = 0) AND (Mus.ISFOLDER = 2) AND (
                                                    (Mus.PARENTID = @Musor) OR Mus.ID in ((select mus1.id from sc84 as Mus1 (NOLOCK) where Mus1.ismark = 0 and Mus1.isfolder = 2 and Mus1.parentid in (select Mus2.id from sc84 as Mus2 (NOLOCK) where Mus2.parentid = @Musor)) UNION
                                                    (select mus3.id from sc84 as Mus3 (NOLOCK) where Mus3.parentid in (select Mus4.id from sc84 as Mus4 (NOLOCK) where Mus4.parentid in (select Mus5.id from sc84 as Mus5 (NOLOCK) where Mus5.parentid = @Musor))))
                                                )";
                    if ((listNomk.Count > 0) && (NomInSelect == 0))
                    {
                        queryString = queryString + @" UNION
                                            SELECT NotInclude.ID FROM SC84 as NotInclude (NOLOCK) WHERE (NotInclude.ISMARK = 0) AND (NotInclude.ISFOLDER = 2) AND 
                                                    (
                                                    NotInclude.PARENTID in ({1}) OR
                                                    NotInclude.ID in ({1}) OR
                                                    NotInclude.ID in (
                                                                        (select NotInclude1.id from sc84 as NotInclude1 (NOLOCK) where NotInclude1.ismark = 0 and NotInclude1.isfolder = 2 and NotInclude1.parentid in (select NotInclude2.id from sc84 as NotInclude2 (NOLOCK) where NotInclude2.parentid in ({1})) UNION
                                                                        (select NotInclude3.id from sc84 as NotInclude3 (NOLOCK) where NotInclude3.parentid in (select NotInclude4.id from sc84 as NotInclude4 (NOLOCK) where NotInclude4.parentid in (select NotInclude5.id from sc84 as NotInclude5 (NOLOCK) where NotInclude5.parentid in ({1})))))
                                                                     )
                                                    )
                    ";
                    }
                    queryString = queryString + @"
                                            ))";
                    if ((listNomk.Count > 0) && (NomInSelect == 1))
                    {
                        queryString = queryString + @" and sp.id in 
                                                    (
                                                        SELECT NotInclude.ID FROM SC84 as NotInclude (NOLOCK) WHERE (NotInclude.ISMARK = 0) AND (NotInclude.ISFOLDER = 2) AND 
                                                                (
                                                                NotInclude.PARENTID in ({1}) OR
                                                                NotInclude.ID in ({1}) OR
                                                                NotInclude.ID in (
                                                                                    (select NotInclude1.id from sc84 as NotInclude1 (NOLOCK) where NotInclude1.ismark = 0 and NotInclude1.isfolder = 2 and NotInclude1.parentid in (select NotInclude2.id from sc84 as NotInclude2 (NOLOCK) where NotInclude2.parentid in ({1})) UNION
                                                                                    (select NotInclude3.id from sc84 as NotInclude3 (NOLOCK) where NotInclude3.parentid in (select NotInclude4.id from sc84 as NotInclude4 (NOLOCK) where NotInclude4.parentid in (select NotInclude5.id from sc84 as NotInclude5 (NOLOCK) where NotInclude5.parentid in ({1})))))
                                                                                 )
                                                                )
                                                    )
                    ";
                    }
                    if (listBrend.Count > 0)
                    {
                        queryString = queryString + " and sp.SP8842 ";
                        if (BrendInSelect == 0)
                            queryString = queryString + "not ";
                        queryString = queryString + "in ({2})";
                    }
                    queryString += " inner join SC75 as SpEd (NOLOCK) on Sp.SP94 = SpEd.id";
                }
                else //по всему справочнику
                {
                    queryString = queryString + @"
                                        From sc84 as Sp (NOLOCK)
                                        inner join SC75 as SpEd (NOLOCK) on Sp.SP94 = SpEd.id
                                        left join 
                                        (
                                            select nomk, Sum(kolvo) as Ostatok, Sum(kolvoPrih) as OstatokPrihod, Sum(kolPartii) as OstatokPartii, Sum(summaPartii) as SummaPartii from 
                                            (
                                            select sp408 as nomk, sp411 as kolvo, 0 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from rg405 (NOLOCK)
                                            where period = @Period and SP418 in (select sc55.id from sc55 (NOLOCK) where sc55.code in ('S0010','S0007','S0013','S0002','S0004','S0006','S0005','D0002','D0001'))
                                            Union all
                                            select SP8698 as nomk, SP8701 as kolvo, 0 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from RG8696 (NOLOCK)
                                            where period = @Period and SP8699 in (select sc55.id from sc55 (NOLOCK) where sc55.code in ('S0010','S0007','S0013','S0002','S0004','S0006','S0005','D0002','D0001'))
                                            Union all
                                            select SP466 as nomk, 0 as kolvo, SP4471 as kolvoPrih, 0 as kolPartii, 0 as summaPartii from RG464 (NOLOCK)
                                            where period = @Period";
                    if (NeedToCalcCost)
                        queryString += @"
                                            Union all
                                            select SP331 as nomk, 0 as kolvo, 0 as kolvoPrih, SP342 as kolPartii, SP421 as summaPartii from RG328 (NOLOCK) 
                                            where period = @Period and SP4061 in ('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + "')";
                    queryString += @"
                                            ) x
                                            group by nomk
                                        ) as R on (Sp.id = R.nomk) ";

                    where = @"WHERE (Sp.IsMark = 0) and (Sp.IsFolder = 2) and (Sp.SP5066 = 0) " + (Top == 1 ? "and (Sp.SP13277 = 1) " : "") + @"and (sp.id not in (
                                            SELECT Mus.ID FROM SC84 as Mus (NOLOCK) WHERE (Mus.ISMARK = 0) AND (Mus.ISFOLDER = 2) AND (
                                                    (Mus.PARENTID = @Musor) OR Mus.ID in ((select mus1.id from sc84 as Mus1 (NOLOCK) where Mus1.ismark = 0 and Mus1.isfolder = 2 and Mus1.parentid in (select Mus2.id from sc84 as Mus2 (NOLOCK) where Mus2.parentid = @Musor)) UNION
                                                    (select mus3.id from sc84 as Mus3 (NOLOCK) where Mus3.parentid in (select Mus4.id from sc84 as Mus4 (NOLOCK) where Mus4.parentid in (select Mus5.id from sc84 as Mus5 (NOLOCK) where Mus5.parentid = @Musor))))
                                                )";
                    if ((listNomk.Count > 0) && (NomInSelect == 0))
                    {
                        where = where + @" UNION
                                            SELECT NotInclude.ID FROM SC84 as NotInclude (NOLOCK) WHERE (NotInclude.ISMARK = 0) AND (NotInclude.ISFOLDER = 2) AND 
                                                    (
                                                    NotInclude.PARENTID in ({1}) OR
                                                    NotInclude.ID in ({1}) OR
                                                    NotInclude.ID in (
                                                                        (select NotInclude1.id from sc84 as NotInclude1 (NOLOCK) where NotInclude1.ismark = 0 and NotInclude1.isfolder = 2 and NotInclude1.parentid in (select NotInclude2.id from sc84 as NotInclude2 (NOLOCK) where NotInclude2.parentid in ({1})) UNION
                                                                        (select NotInclude3.id from sc84 as NotInclude3 (NOLOCK) where NotInclude3.parentid in (select NotInclude4.id from sc84 as NotInclude4 (NOLOCK) where NotInclude4.parentid in (select NotInclude5.id from sc84 as NotInclude5 (NOLOCK) where NotInclude5.parentid in ({1})))))
                                                                     )
                                                    )
                    ";
                    }
                    where = where + @"
                                            ))";
                    if ((listNomk.Count > 0) && (NomInSelect == 1))
                    {
                        where = where + @" and sp.id in 
                                                    (
                                                        SELECT NotInclude.ID FROM SC84 as NotInclude (NOLOCK) WHERE (NotInclude.ISMARK = 0) AND (NotInclude.ISFOLDER = 2) AND 
                                                                (
                                                                NotInclude.PARENTID in ({1}) OR
                                                                NotInclude.ID in ({1}) OR
                                                                NotInclude.ID in (
                                                                                    (select NotInclude1.id from sc84 as NotInclude1 (NOLOCK) where NotInclude1.ismark = 0 and NotInclude1.isfolder = 2 and NotInclude1.parentid in (select NotInclude2.id from sc84 as NotInclude2 (NOLOCK) where NotInclude2.parentid in ({1})) UNION
                                                                                    (select NotInclude3.id from sc84 as NotInclude3 (NOLOCK) where NotInclude3.parentid in (select NotInclude4.id from sc84 as NotInclude4 (NOLOCK) where NotInclude4.parentid in (select NotInclude5.id from sc84 as NotInclude5 (NOLOCK) where NotInclude5.parentid in ({1})))))
                                                                                 )
                                                                )
                                                    )
                    ";
                    }
                    if (listBrend.Count > 0)
                    {
                        where = where + " and (sp.SP8842 ";
                        if (BrendInSelect == 0)
                            where = where + "not ";
                        where = where + "in ({2}))";
                    }

                }
                queryString = queryString + @"
                                        left join SC8840 as SpBrend (NOLOCK) on (Sp.SP8842 = SpBrend.ID)";
                if (NeedToCalcCost)
                {
                    queryString = queryString + @"
                                        left join SC319 as SpCostZakup (NOLOCK) on ((Sp.ID = SpCostZakup.PARENTEXT) AND SpCostZakup.SP327 = @CostZakupType and (SpCostZakup.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstZakup (NOLOCK) ON SpCostZakup.ID = TabConstZakup.OBJID AND TabConstZakup.ID = 324 AND 
                                            TabConstZakup.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstZakup.OBJID and id = TabConstZakup.id
                                                               order by date desc, time desc)
                    ";
                }
                if ((opt == 1) || (osob == 1))
                {
                    queryString = queryString + @"
                                        left join SC319 as SpCost (NOLOCK) on ((Sp.ID = SpCost.PARENTEXT) AND SpCost.SP327 = @CostType and (SpCost.IsMark = 0))
                                        left join sc219 (NOLOCK) on sc219.id = '     PD ' and sc219.IsMark = 0
                                        LEFT OUTER JOIN _1SCONST TabConst (NOLOCK) ON SpCost.ID = TabConst.OBJID AND TabConst.ID = 324 AND 
                                            TabConst.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst.OBJID and id = TabConst.id
                                                               order by date desc, time desc)
                    ";
                }
                if (rozn == 1)
                {
                    queryString = queryString + @"
                                        left join SC319 as SpCostRoz (NOLOCK) on ((Sp.ID = SpCostRoz.PARENTEXT) AND SpCostRoz.SP327 = @CostRoznType and (SpCostRoz.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstRoz (NOLOCK) ON SpCostRoz.ID = TabConstRoz.OBJID AND TabConstRoz.ID = 324 AND 
                                            TabConstRoz.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstRoz.OBJID and id = TabConstRoz.id
                                                                  order by date desc, time desc)
                    ";
                }
                if (spec == 1)
                {
                    queryString = queryString + @"
                                        left join SC319 as SpCostSP (NOLOCK) on ((Sp.ID = SpCostSP.PARENTEXT) and (SpCostSP.SP327 = @CostSPType) and (SpCostSP.IsMark = 0))
                                        left join SC8904 as SpSklCost (NOLOCK) on ((SpCostSP.ID = SpSklCost.PARENTEXT) and (SpSklCost.SP8901 = @CostSkl) and (SpSklCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst4 (NOLOCK) ON SpSklCost.ID = TabConst4.OBJID AND TabConst4.ID = 8902 AND 
                                            TabConst4.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst4.OBJID and id = TabConst4.id
                                                                order by date desc, time desc)
                    ";
                    if (rozn == 1)
                        queryString = queryString + @"
                                        left join SC319 as SpCostSPR (NOLOCK) on ((Sp.ID = SpCostSPR.PARENTEXT) and (SpCostSPR.SP327 = @CostSPType) and (SpCostSPR.IsMark = 0))
                                        left join SC8904 as SpSklCostR (NOLOCK) on ((SpCostSPR.ID = SpSklCostR.PARENTEXT) and (SpSklCostR.SP8901 = @CostSklR) and (SpSklCostR.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst5 (NOLOCK) ON SpSklCostR.ID = TabConst5.OBJID AND TabConst5.ID = 8902 AND 
                                            TabConst5.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst5.OBJID and id = TabConst5.id
                                                                order by date desc, time desc)
                    ";
                }
                queryString = queryString + @"
                                        LEFT OUTER JOIN sc84 As Gr3 (NOLOCK) ON Sp.PARENTID = Gr3.ID
                                        LEFT OUTER JOIN sc84 As Gr2 (NOLOCK) ON Gr3.PARENTID = Gr2.ID
                                        LEFT OUTER JOIN sc84 As Gr1 (NOLOCK) ON Gr2.PARENTID = Gr1.ID ";

                for (int i = 0; i < listSklad.Count; i++)
                {
                    queryString = queryString + @"left join (
                                                    select nomk, (Sum(kolvo) - Sum(rezerv)) as ostatok from 
                                                    (
                                                    select sp408 as nomk, sp411 as kolvo, 0 as rezerv from rg405 (NOLOCK) 
                                                        where period = @Period and sp418 = '" + listSklad[i] + @"'
                                                    Union all
                                                    select sp4477 as nomk, 0 as kolvo, sp4479 as rezerv from RG4480 (NOLOCK)
                                                        where period = @Period and SP4476 = '" + listSklad[i] + @"'
                                                    Union all
                                                    select SP8698 as nomk, SP8701 as kolvo, 0 as rezerv from RG8696 (NOLOCK)
                                                    where period = @Period and SP8699 = '" + listSklad[i] + @"'
                                                    ) xx group by nomk)
                                                    as R_Skl" + i.ToString() + " on Sp.ID = R_Skl" + i.ToString() + ".nomk ";
                }

                queryString = queryString + where;
                //
                //                                        ORDER BY Sp.Descr;";

                string[] paramSkl = listSklad.Select(x => "'" + x + "'").ToArray();
                string[] paramNom = listNomk.Select(x => "'" + x + "'").ToArray();
                string[] paramBr = listBrend.Select(x => "'" + x + "'").ToArray();
                if (NeedToCalcCost)
                {
                    string[] paramCustomer = (new List<string>(CustomerListStr.Split(','))).Select(x => "'" + x + "'").ToArray();
                    queryString = string.Format(queryString, string.Join(",", paramSkl), string.Join(",", paramNom), string.Join(",", paramBr), string.Join(",", paramCustomer));
                }
                else
                    queryString = string.Format(queryString, string.Join(",", paramSkl), string.Join(",", paramNom), string.Join(",", paramBr));

                SqlCommand command = new SqlCommand(queryString, connection);
                if (NeedToCalcCost)
                    command.Parameters.AddWithValue("@CostZakupType", CostType.Zak);
                if ((opt == 1) || (osob == 1))
                {
                    command.Parameters.AddWithValue("@CostType", CostType.Opt);
                }
                if (rozn == 1)
                    command.Parameters.AddWithValue("@CostRoznType", CostType.Rozn);
                if (spec == 1)
                {
                    command.Parameters.AddWithValue("@CostSPType", CostType.SP);
                    command.Parameters.AddWithValue("@CostSkl", Sklad.Ekran);
                    if (rozn == 1)
                        command.Parameters.AddWithValue("@CostSklR", Sklad.Gastello);
                }
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                command.Parameters.AddWithValue("@Musor", Nomenklatura.Musor);
                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                command.Parameters["@Period"].Value = startDate;
                command.CommandTimeout = 500;

                connection.Open();

                SqlDataReader reader = command.ExecuteReader();
                List<Price> rs = new List<Price>();
                List<Price> Sorted_rs = new List<Price>();
                List<string> UsedIDs = new List<string>();
                double porog = 0;
                double porogBr = 0;
                int OtsrDney = 0;
                Customer cust = new Customer();
                try
                {
                    if (NeedToCalcCost)
                    {
                        //порог себестоимости
                        while (reader.Read())
                        {
                            porog = Convert.ToDouble(reader.GetValue(0), CultureInfo.InvariantCulture);
                        }
                        if (reader.NextResult())
                            while (reader.Read())
                            {
                                cust.Id = reader.GetString(0);
                                cust.Code = reader.GetString(1);
                                cust.Name = reader.GetString(2);
                                cust.MainDogovorId = reader.GetString(3);
                                cust.TypeCostId = reader.GetString(4);
                                cust.SkidkaId = reader.GetString(5);
                                cust.GroupId = reader.GetString(6);
                                cust.Export = Convert.ToBoolean(reader.GetValue(7), CultureInfo.InvariantCulture);
                                cust.KolonkaSkidki = reader.GetString(8);
                                cust.SkidkaValue = Convert.ToDouble(reader.GetValue(9), CultureInfo.InvariantCulture);
                                OtsrDney = Convert.ToInt32(reader.GetValue(10), CultureInfo.InvariantCulture);
                            }
                        //колонки скидок
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                switch (cust.KolonkaSkidki.Trim())
                                {
                                    case "93R":
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture) });
                                        break;
                                    case "93S":
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture) });
                                        break;
                                    case "93T":
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture) });
                                        break;
                                    case "93U":
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(4), CultureInfo.InvariantCulture) });
                                        break;
                                    case "93V":
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(5), CultureInfo.InvariantCulture) });
                                        break;
                                    default:
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = 0 });
                                        break;
                                }
                            }
                        }
                        //доп условия клиента
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                string customer = reader.GetString(0);
                                string proiz = reader.GetString(1);
                                double skidka = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture);

                                if (IsMasterskaya == 0)
                                {
                                    if (cust.proizSkidka.Any(x => x.ProizId == proiz))
                                        foreach (ProizSkidka ps in cust.proizSkidka.Where(x => x.ProizId == proiz))
                                        {
                                            ps.Skidka = skidka;
                                        }
                                    else
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = proiz, Skidka = skidka });
                                }
                                else
                                {
                                    if (cust.proizSkidkaM.Any(x => x.ProizId == proiz))
                                        foreach (ProizSkidka ps in cust.proizSkidkaM.Where(x => x.ProizId == proiz))
                                        {
                                            ps.Skidka = skidka;
                                        }
                                    else
                                        cust.proizSkidkaM.Add(new ProizSkidka { ProizId = proiz, Skidka = skidka });
                                }
                            }
                        }
                        reader.NextResult();
                    }
                    List<string> skList = new List<string>();
                    while (reader.Read())
                    {
                        skList.Add(reader.GetString(0).Trim());
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            double CostOpt = 0;
                            double CostRoz = 0;
                            double CostSuperOpt = 0;
                            double CostSuperRoz = 0;
                            string _ProizId = "";
                            if ((opt == 1) || (osob == 1))
                            {
                                if (!reader.IsDBNull(16 + listSklad.Count))
                                    CostOpt = Convert.ToDouble(reader.GetValue(16 + listSklad.Count), CultureInfo.InvariantCulture);
                            }
                            if (spec == 1)
                            {
                                int NumSpec = 16 + listSklad.Count;
                                if ((opt == 1) || (osob == 1))
                                    NumSpec = NumSpec + 3;
                                if (!reader.IsDBNull(NumSpec))
                                    CostSuperOpt = Convert.ToDouble(reader.GetValue(NumSpec), CultureInfo.InvariantCulture);
                                if (rozn == 1)
                                    CostSuperRoz = Convert.ToDouble(reader.GetValue(NumSpec + 1), CultureInfo.InvariantCulture);

                            }
                            if (rozn == 1)
                            {
                                int NumRozn = 16 + listSklad.Count;
                                if ((opt == 1) || (osob == 1))
                                    NumRozn = NumRozn + 3;
                                if (spec == 1)
                                    NumRozn = NumRozn + 2;
                                if (!reader.IsDBNull(NumRozn))
                                    CostRoz = Convert.ToDouble(reader.GetValue(NumRozn), CultureInfo.InvariantCulture);
                            }
                            if (onlyOst == 1)
                            {
                                double Ostatok = 0;
                                if (!reader.IsDBNull(11))
                                    Ostatok += Convert.ToDouble(reader.GetValue(11), CultureInfo.InvariantCulture);
                                if (!reader.IsDBNull(12))
                                    Ostatok += Convert.ToDouble(reader.GetValue(12), CultureInfo.InvariantCulture);
                                if (Ostatok <= 0)
                                {
                                    //не будем выводить
                                    CostOpt = 0;
                                    CostRoz = 0;
                                    CostSuperOpt = 0;
                                    CostSuperRoz = 0;
                                }
                            }
                            if (CostOpt + CostRoz + CostSuperOpt + CostSuperRoz > 0)
                            {

                                Price pr = new Price();
                                if (!reader.IsDBNull(0))
                                {
                                    Price pr1 = new Price();
                                    pr1.level = 1;
                                    pr1.ID = reader.GetString(0).Trim();
                                    pr1.Nomenklatura = reader.GetString(1).Trim();
                                    pr1.IsFolder = 1;
                                    if (!UsedIDs.Contains(pr1.ID))
                                    {
                                        rs.Add(pr1);
                                        UsedIDs.Add(pr1.ID);
                                    }

                                    pr.level = 4;
                                    pr.ParentID1 = reader.GetString(0).Trim();
                                }
                                if (!reader.IsDBNull(2))
                                {
                                    Price pr2 = new Price();
                                    pr2.ID = reader.GetString(2).Trim();
                                    pr2.Nomenklatura = reader.GetString(3).Trim();
                                    pr2.IsFolder = 1;
                                    if (pr.level == 4)
                                    {
                                        pr.ParentID2 = reader.GetString(2).Trim();

                                        pr2.level = 2;
                                        pr2.ParentID1 = pr.ParentID1;
                                        pr2.ParentID = pr.ParentID1;
                                    }
                                    else
                                    {
                                        pr.level = 3;
                                        pr.ParentID1 = reader.GetString(2).Trim();

                                        pr2.level = 1;
                                    }
                                    if (!UsedIDs.Contains(pr2.ID))
                                    {
                                        rs.Add(pr2);
                                        UsedIDs.Add(pr2.ID);
                                    }
                                }
                                if (!reader.IsDBNull(4))
                                {
                                    Price pr3 = new Price();
                                    pr3.ID = reader.GetString(4).Trim();
                                    pr3.Nomenklatura = reader.GetString(5).Trim();
                                    pr3.IsFolder = 1;
                                    if (pr.level == 4)
                                    {
                                        pr.ParentID3 = reader.GetString(4).Trim();

                                        pr3.level = 3;
                                        pr3.ParentID1 = pr.ParentID1;
                                        pr3.ParentID2 = pr.ParentID2;
                                        pr3.ParentID = pr.ParentID2;
                                    }
                                    else if (pr.level == 3)
                                    {
                                        pr.ParentID2 = reader.GetString(4).Trim();

                                        pr3.level = 2;
                                        pr3.ParentID1 = pr.ParentID1;
                                        pr3.ParentID = pr.ParentID1;
                                    }
                                    else
                                    {
                                        pr.level = 2;
                                        pr.ParentID1 = reader.GetString(4).Trim();

                                        pr3.level = 1;
                                    }
                                    if (!UsedIDs.Contains(pr3.ID))
                                    {
                                        rs.Add(pr3);
                                        UsedIDs.Add(pr3.ID);
                                    }
                                }
                                if (pr.level == 0)
                                    pr.level = 1;
                                if (!reader.IsDBNull(6))
                                    pr.ID = reader.GetString(6).Trim();
                                if (!reader.IsDBNull(4))
                                    pr.ParentID = reader.GetString(4).Trim();
                                if (!reader.IsDBNull(7))
                                    pr.Articul = reader.GetString(7).Trim();
                                //LogMessageToFile(pr.Articul);
                                if (!reader.IsDBNull(8))
                                    pr.Nomenklatura = reader.GetString(8).Trim();
                                if (!reader.IsDBNull(9))
                                    pr.Brend = reader.GetString(9).Trim();
                                if (!reader.IsDBNull(10))
                                    pr.Harakter = reader.GetString(10).Trim();
                                pr.NeNadoZakup = Convert.ToBoolean(reader["SP10397"]);
                                if (!reader.IsDBNull(11))
                                    pr.Ostatok = Convert.ToDouble(reader.GetValue(11), CultureInfo.InvariantCulture);
                                if (!reader.IsDBNull(12))
                                    pr.OstatokPrihod = Convert.ToDouble(reader.GetValue(12), CultureInfo.InvariantCulture);
                                double kol = 0;
                                double summa = 0;
                                if (!reader.IsDBNull(13))
                                    kol = Convert.ToDouble(reader.GetValue(13), CultureInfo.InvariantCulture);
                                if (!reader.IsDBNull(14))
                                    summa = Convert.ToDouble(reader.GetValue(14), CultureInfo.InvariantCulture);
                                double koff = Convert.ToDouble(reader["koff"], CultureInfo.InvariantCulture);
                                double sebestoim = 0;
                                if (kol > 0)
                                    sebestoim = summa / Math.Round(kol / koff, 3);
                                _ProizId = reader.GetString(15);
                                pr.OstatokSkl = new List<double>();
                                for (int i = 0; i < listSklad.Count; i++)
                                {
                                    pr.OstatokSkl.Add(Convert.ToDouble(reader.GetValue(i + 16), CultureInfo.InvariantCulture));
                                }
                                if ((opt == 1) || (osob == 1))
                                {
                                    if (!reader.IsDBNull(16 + listSklad.Count))
                                        pr.Cost = Convert.ToDouble(reader.GetValue(16 + listSklad.Count), CultureInfo.InvariantCulture);
                                    double Baza = 0;
                                    if (!reader.IsDBNull(16 + listSklad.Count + 1))
                                        Baza = Convert.ToDouble(reader.GetValue(16 + listSklad.Count + 1), CultureInfo.InvariantCulture);
                                    if (!reader.IsDBNull(16 + listSklad.Count + 2))
                                        Baza += Convert.ToDouble(reader.GetValue(16 + listSklad.Count + 2), CultureInfo.InvariantCulture);
                                    pr.CostOsob = Math.Round(pr.Cost * (100 + Baza) / 100, 2);
                                    porogBr = Convert.ToDouble(reader.GetValue(16 + listSklad.Count + 2), CultureInfo.InvariantCulture);
                                }
                                if (spec == 1)
                                {
                                    int NumSpec = 16 + listSklad.Count;
                                    if ((opt == 1) || (osob == 1))
                                        NumSpec = NumSpec + 3;
                                    if (!reader.IsDBNull(NumSpec))
                                        pr.SuperCost = Convert.ToDouble(reader.GetValue(NumSpec), CultureInfo.InvariantCulture);
                                    if (rozn == 1)
                                        pr.SuperCostR = Convert.ToDouble(reader.GetValue(NumSpec + 1), CultureInfo.InvariantCulture);

                                }
                                if (rozn == 1)
                                {
                                    int NumRozn = 16 + listSklad.Count;
                                    if ((opt == 1) || (osob == 1))
                                        NumRozn = NumRozn + 3;
                                    if (spec == 1)
                                        NumRozn = NumRozn + 2;
                                    //NumRozn = NumRozn + spec;
                                    if (!reader.IsDBNull(NumRozn))
                                        pr.CostRozn = Convert.ToDouble(reader.GetValue(NumRozn), CultureInfo.InvariantCulture);
                                }
                                if (NeedToCalcCost)
                                {
                                    Tovar tov = new Tovar();
                                    tov.Costs = new Costs { Zak = Convert.ToDouble(reader["costZakup"], CultureInfo.InvariantCulture), Opt = pr.Cost, Osob = pr.CostOsob, Rozn = pr.CostRozn, SP = pr.SuperCost, SPR = pr.SuperCostR };
                                    tov.ProizId = _ProizId;
                                    tov.sebestoim = sebestoim;
                                    tov.PorogSebestoim = porog + porogBr;
                                    pr.CostClient = CalculateSumma(tov, cust, 1, 0, (IsMasterskaya == 1));
                                }
                                pr.IsFolder = 2;
                                rs.Add(pr);
                            }
                        }
                    }

                    rs = rs.OrderBy(x => x.ParentID1).ThenBy(x => x.Nomenklatura).ToList();
                    SetTree(Sorted_rs, rs, null);

                    PrList.ExcelFile = PriceXLS.CreateXLS(listUsedNomk, skList, harakter, opt, osob, spec, rozn, zakaz, (NeedToCalcCost ? 1 : 0), Sorted_rs.ToArray(), HttpContext.Current.Server.MapPath("~/bin/") + "PriceTemplate." + ext, Address, OtsrDney, ostPrice);

                    PrList.ReplyState.State = "Success";
                    PrList.ReplyState.Description = "Информация успешно извлечена из БД";
                }
                catch (Exception e)
                {
                    PrList.ReplyState.State = "Error";
                    PrList.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }
            return PrList;
        }

        private string SprInclude(string typeName, string Name, string ParamName, bool In = true, bool OneLevel = false)
        {
            string q;
            if (OneLevel)
            {
                q = Name + ".id " + (In ? "" : "not") + " in (" + ParamName + ") ";
            }
            else
            {
                if (!In)
                {
                    q = "(" + Name + ".id not in (" + ParamName + ") and ";
                    q = q + Name + ".parentid not in (" + ParamName + ") and ";
                    q = q + Name + ".id not in (select sp3.id from " + typeName + " as sp3 (NOLOCK) where sp3.ismark = 0 and sp3.isfolder = 2 and sp3.parentid in (select sp2.id from " + typeName + " as sp2 (NOLOCK) where sp2.ismark = 0 and sp2.isfolder = 1 and sp2.parentid in (" + ParamName + "))) and ";
                    q = q + Name + ".parentid not in (select sp3.id from " + typeName + " as sp3 (NOLOCK) where sp3.ismark = 0 and sp3.isfolder = 1 and sp3.parentid in (select sp2.id from " + typeName + " as sp2 (NOLOCK) where sp2.ismark = 0 and sp2.isfolder = 1 and sp2.parentid in (" + ParamName + ")))) ";
                }
                else
                {
                    q = "(" + Name + ".id in (" + ParamName + ") or ";
                    q = q + Name + ".parentid in (" + ParamName + ") or ";
                    q = q + Name + ".id in (select sp3.id from " + typeName + " as sp3 (NOLOCK) where sp3.ismark = 0 and sp3.isfolder = 2 and sp3.parentid in (select sp2.id from " + typeName + " as sp2 (NOLOCK) where sp2.ismark = 0 and sp2.isfolder = 1 and sp2.parentid in (" + ParamName + "))) or ";
                    q = q + Name + ".parentid in (select sp3.id from " + typeName + " as sp3 (NOLOCK) where sp3.ismark = 0 and sp3.isfolder = 1 and sp3.parentid in (select sp2.id from " + typeName + " as sp2 (NOLOCK) where sp2.ismark = 0 and sp2.isfolder = 1 and sp2.parentid in (" + ParamName + ")))) ";
                }
            }
            return q;
        }

        private string CheckForQuotes(string AString)
        {
            if (String.IsNullOrEmpty(AString))
                return AString;
            string quote = "\"";
            AString = AString.Replace(";", "");
            AString = AString.Replace("=", " ");
            if ((AString[0] == '"') && (AString[AString.Length - 1] == '"'))
                AString = AString.Substring(1, AString.Length - 2);
            if (AString.Contains(quote))
                AString = quote + AString.Replace(quote, quote + quote) + quote;

            return AString;
        }

        private string RemoveLeadingNumbers(string AString)
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

        private string AddReplyString(string AString, string AddString)
        {
            if (!string.IsNullOrEmpty(AString))
                AString = AString + System.Environment.NewLine;
            AString = AString + AddString;
            return AString;
        }

        private Int64 Decode36(string input)
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

        private String Encode36(long input)
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

        private void CalculateCostHelper(List<Tovar> Tovars, List<Customer> Customers)
        {
            DateTime TA = GetDateTA();

            using (SqlConnection connection = new SqlConnection(connDocsStr))
            {
                //себестоимость
                string queryString = @"select SC84.id, SpEd.sp78 as koff,Sum(RG328.SP342) as kol,Sum(RG328.SP421) as summa from RG328 (NOLOCK)
                                        inner join SC84 (NOLOCK) on SC84.id = RG328.SP331
                                        inner join SC75 as SpEd (NOLOCK) on sc84.sp94 = SpEd.id
                                        where RG328.period = convert(varchar(8), @Period, 112) and SC84.code in ({0}) and
                                            RG328.SP4061 in ('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + @"')
                                        group by SC84.id, SpEd.sp78;";
                string[] paramTovar = Tovars.Select(t => "'" + t.Code + "'").Distinct().ToArray();
                queryString = string.Format(queryString, string.Join(",", paramTovar));
                //порог себестоимости
                queryString += @"select value from _1SCONST(NOLOCK) where id = 12572;";
                //колонки скидок
                queryString += @"select distinct SP11784,SP11785,SP11786,SP11787,SP11788,SP11789 from SC11791 (NOLOCK)
                                    where SP11784 in ({0}) and (IsMark = 0);";
                string[] paramProiz = Tovars.Select(f => "'" + f.ProizId + "'").Distinct().ToArray();
                queryString = string.Format(queryString, string.Join(",", paramProiz));
                //доп условия клиента
                queryString += @"select SC9671.parentext, SC9671.SP9667, SC426.SP429 from SC9671 (NOLOCK)
                                    inner join SC426 (NOLOCK) on SC426.id = SC9671.SP9668 
                                    where SC9671.ismark = 0 and SC9671.parentext in ({0}) and SC9671.SP11255 = 1;";
                //доп условия клиента (мастерская)
                queryString += @"select SC12703.parentext, SC12703.SP12698, SC426.SP429 from SC12703 (NOLOCK)
                                    inner join SC426 (NOLOCK) on SC426.id = SC12703.SP12699 
                                    where SC12703.ismark = 0 and SC12703.parentext in ({0}) and SC12703.SP12701 = 1;";
                string[] paramCust = Customers.Select(f => "'" + f.Id + "'").Distinct().ToArray();
                queryString = string.Format(queryString, string.Join(",", paramCust));
                //бесплатная отсрочка 
                queryString += @"select SC172.id, SC12695.SP12693 from SC12695 (NOLOCK)
                                    inner join SC172 (NOLOCK) on SC172.SP9631 = SC12695.parentext
                                    where SC12695.parentext in ({0}) and SC12695.ismark = 0;";
                //бесплатная отсрочка (мастерская)
                queryString += @"select SC172.id, SC12709.SP12707 from SC12709 (NOLOCK)
                                    inner join SC172 (NOLOCK) on SC172.SP9631 = SC12709.parentext
                                    where SC12709.parentext in ({0}) and SC12709.ismark = 0;";
                string[] paramGroupCust = Customers.Select(x => "'" + x.GroupId + "'").Distinct().ToArray();
                queryString = string.Format(queryString, string.Join(",", paramGroupCust));
                //цены номенклатуры
                queryString += @"select sc84.id, ISNULL(TabConst.VALUE, 0), IsNull(sc219.sp221, 0), IsNull(SpBrend.SP11668, 0), IsNull(SC426.SP429, 0), ISNULL(TabConstRoz.VALUE, 0), ISNULL(TabConstSP.VALUE, 0), ISNULL(TabConstSPR.VALUE, 0), ISNULL(TabConstZakup.VALUE, 0)
                                    from SC84 (NOLOCK)
                                        left join SC8840 as SpBrend (NOLOCK) on (sc84.SP8842 = SpBrend.ID)
                                        left join SC426 (NOLOCK) on SC426.id = SpBrend.SP11254 
                                        left join SC319 as SpCostZakup (NOLOCK) on ((sc84.ID = SpCostZakup.PARENTEXT) AND SpCostZakup.SP327 = @CostZakupType and (SpCostZakup.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstZakup (NOLOCK) ON SpCostZakup.ID = TabConstZakup.OBJID AND TabConstZakup.ID = 324 AND 
                                            TabConstZakup.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstZakup.OBJID and id = TabConstZakup.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCost (NOLOCK) on ((sc84.ID = SpCost.PARENTEXT) AND SpCost.SP327 = @CostType and (SpCost.IsMark = 0))
                                        left join sc219 (NOLOCK) on sc219.id = '     PD ' and sc219.IsMark = 0
                                        LEFT OUTER JOIN _1SCONST TabConst (NOLOCK) ON SpCost.ID = TabConst.OBJID AND TabConst.ID = 324 AND 
                                            TabConst.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst.OBJID and id = TabConst.id
                                                               order by date desc, time desc)
                                        left join SC319 as SpCostRoz (NOLOCK) on ((sc84.ID = SpCostRoz.PARENTEXT) AND SpCostRoz.SP327 = @CostRoznType and (SpCostRoz.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstRoz (NOLOCK) ON SpCostRoz.ID = TabConstRoz.OBJID AND TabConstRoz.ID = 324 AND 
                                            TabConstRoz.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstRoz.OBJID and id = TabConstRoz.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCostSP (NOLOCK) on ((sc84.ID = SpCostSP.PARENTEXT) and (SpCostSP.SP327 = @CostSPType) and (SpCostSP.IsMark = 0))
                                        left join SC8904 as SpSklCost (NOLOCK) on ((SpCostSP.ID = SpSklCost.PARENTEXT) and (SpSklCost.SP8901 = @CostSkl) and (SpSklCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstSP (NOLOCK) ON SpSklCost.ID = TabConstSP.OBJID AND TabConstSP.ID = 8902 AND 
                                            TabConstSP.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstSP.OBJID and id = TabConstSP.id
                                                                order by date desc, time desc)
                                        left join SC8904 as SpSklCostR (NOLOCK) on ((SpCostSP.ID = SpSklCostR.PARENTEXT) and (SpSklCostR.SP8901 = @CostSklR) and (SpSklCostR.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstSPR (NOLOCK) ON SpSklCostR.ID = TabConstSPR.OBJID AND TabConstSPR.ID = 8902 AND 
                                            TabConstSPR.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstSPR.OBJID and id = TabConstSPR.id
                                                                order by date desc, time desc)
                                        where sc84.id in ({0});";
                paramTovar = Tovars.Select(t => "'" + t.Id + "'").Distinct().ToArray();
                queryString = string.Format(queryString, string.Join(",", paramTovar));


                using (SqlCommand command = new SqlCommand(queryString, connection))
                {
                    command.Parameters.Add("@Period", SqlDbType.DateTime);
                    command.Parameters["@Period"].Value = new DateTime(TA.Year, TA.Month, 1);
                    command.Parameters.AddWithValue("@CostZakupType", CostType.Zak);
                    command.Parameters.AddWithValue("@CostType", CostType.Opt);
                    command.Parameters.AddWithValue("@CostRoznType", CostType.Rozn);
                    command.Parameters.AddWithValue("@CostSPType", CostType.SP);
                    command.Parameters.AddWithValue("@CostSkl", Sklad.Ekran);
                    command.Parameters.AddWithValue("@CostSklR", Sklad.Gastello);

                    connection.Open();

                    double porog = 0;
                    SqlDataReader reader = command.ExecuteReader();
                    try
                    {
                        //себестоимость 
                        while (reader.Read())
                        {
                            string NomId = reader.GetString(0);
                            foreach (Tovar tov in Tovars.Where(x => x.Id == NomId))
                            {
                                double kol = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture);
                                double summa = Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture);
                                double koff = Convert.ToDouble(reader["koff"], CultureInfo.InvariantCulture);
                                if (kol > 0)
                                    tov.sebestoim = summa / Math.Round(kol / koff, 3);
                            }
                        }
                        //порог себестоимости
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                porog = Convert.ToDouble(reader.GetValue(0), CultureInfo.InvariantCulture);
                            }
                        }
                        //колонки скидок
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                foreach (Customer cust in Customers)
                                {
                                    switch (cust.KolonkaSkidki.Trim())
                                    {
                                        case "93R":
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture) });
                                            break;
                                        case "93S":
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture) });
                                            break;
                                        case "93T":
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture) });
                                            break;
                                        case "93U":
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(4), CultureInfo.InvariantCulture) });
                                            break;
                                        case "93V":
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = Convert.ToDouble(reader.GetValue(5), CultureInfo.InvariantCulture) });
                                            break;
                                        default:
                                            cust.proizSkidka.Add(new ProizSkidka { ProizId = reader.GetString(0), Skidka = 0 });
                                            break;

                                    }
                                }
                            }
                        }
                        //доп условия клиента
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                string customer = reader.GetString(0);
                                string proiz = reader.GetString(1);
                                double skidka = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture);

                                foreach (Customer cust in Customers.Where(x => x.Id == customer))
                                {

                                    bool containItem = cust.proizSkidka.Any(x => x.ProizId == proiz);
                                    if (containItem)
                                        foreach (ProizSkidka ps in cust.proizSkidka.Where(x => x.ProizId == proiz))
                                        {
                                            ps.Skidka = skidka;
                                        }
                                    else
                                        cust.proizSkidka.Add(new ProizSkidka { ProizId = proiz, Skidka = skidka });
                                }
                            }
                        }
                        //доп условия клиента (мастерская)
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                string customer = reader.GetString(0);
                                string proiz = reader.GetString(1);
                                double skidka = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture);

                                foreach (Customer cust in Customers.Where(x => x.Id == customer))
                                {

                                    bool containItem = cust.proizSkidkaM.Any(x => x.ProizId == proiz);
                                    if (containItem)
                                        foreach (ProizSkidka ps in cust.proizSkidkaM.Where(x => x.ProizId == proiz))
                                        {
                                            ps.Skidka = skidka;
                                        }
                                    else
                                        cust.proizSkidkaM.Add(new ProizSkidka { ProizId = proiz, Skidka = skidka });
                                }
                            }
                        }
                        //бесплатная отсрочка 
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                string customer = reader.GetString(0);
                                foreach (Customer cust in Customers.Where(x => x.Id == customer))
                                {
                                    cust.proizOtsrochka0.Add(reader.GetString(1));
                                }
                            }
                        }
                        //бесплатная отсрочка (мастерская)
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                string customer = reader.GetString(0);
                                foreach (Customer cust in Customers.Where(x => x.Id == customer))
                                {
                                    cust.proizOtsrochka0M.Add(reader.GetString(1));
                                }
                            }
                        }
                        //цены номенклатуры
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                foreach (Tovar tov in Tovars.Where(x => x.Id == reader.GetString(0)))
                                {
                                    tov.Costs.Zak = Convert.ToDouble(reader.GetValue(8), CultureInfo.InvariantCulture);
                                    tov.Costs.Opt = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                                    tov.Costs.Rozn = Convert.ToDouble(reader.GetValue(5), CultureInfo.InvariantCulture);
                                    tov.Costs.SP = Convert.ToDouble(reader.GetValue(6), CultureInfo.InvariantCulture);
                                    tov.Costs.SPR = Convert.ToDouble(reader.GetValue(7), CultureInfo.InvariantCulture);
                                    double Baza = Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture);
                                    double BazaPr = Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture);
                                    tov.Costs.Osob = Math.Round(tov.Costs.Opt * (100 + Baza + BazaPr) / 100, 2);

                                    tov.PorogSebestoim = porog + BazaPr;

                                    tov.SkidkaVsem = Convert.ToDouble(reader.GetValue(4), CultureInfo.InvariantCulture);
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        reader.Close();
                    }
                }
            }
        }

        private double CalculateSumma(Tovar tov, Customer cust, double Quantity, double SkidkaNovTovar, bool IsMasterskaya)
        {
            double summa = 0;
            double cost = 0;
            double costSpr = 0;
            double sp = tov.Costs.SPR;
            double skidka = cust.SkidkaValue;
            double skidkaVsem = 0;
            bool CheckOsob = false;
            switch (cust.TypeCostId)
            {
                case CostType.Opt:
                    costSpr = tov.Costs.Opt;
                    cost = tov.Costs.Opt;
                    ProizSkidka item;
                    string proizOtsr0;
                    if (IsMasterskaya)
                    {
                        item = cust.proizSkidkaM.FirstOrDefault(x => x.ProizId == tov.ProizId);
                        proizOtsr0 = cust.proizOtsrochka0M.FirstOrDefault(x => x == tov.ProizId);
                    }
                    else
                    {
                        item = cust.proizSkidka.FirstOrDefault(x => x.ProizId == tov.ProizId);
                        proizOtsr0 = cust.proizOtsrochka0.FirstOrDefault(x => x == tov.ProizId);
                    }
                    double itemSkidka = 0;
                    if (item != null)
                        itemSkidka = item.Skidka;
                    cost = cost - (itemSkidka + SkidkaNovTovar) / 100 * cost;
                    sp = tov.Costs.SP;
                    if (!string.IsNullOrEmpty(proizOtsr0))
                        skidka = cust.SkidkaValue;
                    skidkaVsem = tov.SkidkaVsem;
                    break;
                case CostType.Rozn:
                    costSpr = tov.Costs.Rozn;
                    cost = tov.Costs.Rozn;
                    cost = cost - SkidkaNovTovar / 100 * cost;
                    if ((cost - cost * (skidka + skidkaVsem) / 100) < tov.Costs.Osob)
                    {
                        cost = tov.Costs.Osob;
                        CheckOsob = true;
                    }
                    break;
                case CostType.Osob:
                    costSpr = tov.Costs.Osob;
                    cost = tov.Costs.Osob;
                    cost = cost - SkidkaNovTovar / 100 * cost;
                    break;
                default:
                    costSpr = tov.Costs.Rozn;
                    cost = tov.Costs.Rozn;
                    cost = cost - SkidkaNovTovar / 100 * cost;
                    if ((cost - cost * (skidka + skidkaVsem) / 100) < tov.Costs.Osob)
                    {
                        cost = tov.Costs.Osob;
                        CheckOsob = true;
                    }
                    break;
            }
            double PorogValue = Math.Max(tov.sebestoim, tov.Costs.Zak);
            double porog = Math.Min((PorogValue != 0 ? PorogValue * (100 + tov.PorogSebestoim) / 100 : costSpr), costSpr);

            if (porog > cost)
                cost = porog;

            if ((sp > 0) && (sp < cost))
                cost = sp;

            if (cust.Export)
            {
                if (DateTime.Today.Year < 2019)
                    cost = cost - cost * 18 / 118;
                else
                    cost = cost - cost * 20 / 120;
            }

            summa = Quantity * cost;
            if (!CheckOsob)
                summa = summa - summa * (skidka + skidkaVsem) / 100;

            return summa;
        }

        private string GenerateIdDoc(SqlConnection connection, SqlTransaction trans = null)
        {
            string result = "";
            if (connection != null && connection.State == ConnectionState.Open)
            {
                string queryIdDoc = "select max(iddoc) from _1sjourn (NOLOCK)";
                Int64 num10 = 0;
                string num36 = "";
                string num36_pf = "";
                if (trans == null)
                    using (SqlCommand GetMaxIdDoc = new SqlCommand(queryIdDoc, connection))
                    {
                        using (SqlDataReader reader = GetMaxIdDoc.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (reader.Read())
                            {
                                if (!reader.IsDBNull(0))
                                {
                                    num36 = reader.GetString(0);
                                    num36_pf = num36.Substring(6);
                                    num36 = num36.Substring(0, 6).Trim();
                                }
                            }
                        }
                    }
                else
                    using (SqlCommand GetMaxIdDoc = new SqlCommand(queryIdDoc, connection, trans))
                    {
                        using (SqlDataReader reader = GetMaxIdDoc.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (reader.Read())
                            {
                                if (!reader.IsDBNull(0))
                                {
                                    num36 = reader.GetString(0);
                                    num36_pf = num36.Substring(6);
                                    num36 = num36.Substring(0, 6).Trim();
                                }
                            }
                        }
                    }

                if (!string.IsNullOrEmpty(num36))
                    num10 = Decode36(num36);
                num10++;
                result = (Encode36(num10) + num36_pf).PadLeft(9);
            }
            return result;
        }

        private void Fill_1SCRDOC(SqlConnection connection, SqlTransaction trans, int MdId, string ParentVal, string ChildDateTimeIdDoc, string ChildId)
        {
            using (SqlCommand com1SCRDOC_Write = new SqlCommand("_1sp__1SCRDOC_Write", connection, trans) { CommandType = CommandType.StoredProcedure })
            {
                com1SCRDOC_Write.Parameters.AddWithValue("@MdId", MdId);
                com1SCRDOC_Write.Parameters.AddWithValue("@ParentVal", ParentVal);
                com1SCRDOC_Write.Parameters.AddWithValue("@ChildDateTimeIdDoc", ChildDateTimeIdDoc);
                com1SCRDOC_Write.Parameters.AddWithValue("@ChildId", ChildId);
                com1SCRDOC_Write.Parameters.AddWithValue("@Flags", 1);
                com1SCRDOC_Write.ExecuteNonQuery();
            }
        }

        private void RegisterUpdate(SqlConnection connection, SqlTransaction trans, string DbSign, int TypeID, string ObjID, string Deleted)
        {
            using (SqlCommand com1SRegisterUpdate = new SqlCommand("_1sp_RegisterUpdate", connection, trans) { CommandType = CommandType.StoredProcedure })
            {
                com1SRegisterUpdate.Parameters.AddWithValue("@p1", DbSign);
                com1SRegisterUpdate.Parameters.AddWithValue("@p2", TypeID);
                com1SRegisterUpdate.Parameters.AddWithValue("@p3", ObjID);
                com1SRegisterUpdate.Parameters.AddWithValue("@p4", Deleted);
                com1SRegisterUpdate.ExecuteNonQuery();
            }
        }

        private Doc NewDoc(string[] data, string RoznCustomer, string OrderType = "", string OrderNumerator = "", Doc Parent = null)
        {
            if ((data[0] == "0") | (data[0] == "00"))
                return new Doc()
                {
                    Order_Firma = (((data.Length > 6) && (data[6] == "2")) ? Firma.Stin_service : Firma.IP_pavlov),
                    Order_id = data[1],
                    Order_date = DateTime.Now,
                    Order_Customer = (string.IsNullOrEmpty(data[3]) ? RoznCustomer : data[3]),
                    Order_Ref = (data[4].Length > 150 ? data[4].Substring(0, 150) : data[4]),
                    Order_TotalSum = Convert.ToDouble(data[5], CultureInfo.InvariantCulture),
                    Records = new List<DocRecord>()
                };
            else
                return new Doc()
                {
                    Order_type = OrderType,
                    Order_numerator = (string.IsNullOrEmpty(OrderNumerator) ? OrderType : OrderNumerator),
                    Order_parent = Parent,
                    Order_Firma = (((data.Length > 6) && (data[6] == "2")) ? Firma.Stin_service : Firma.IP_pavlov),
                    Order_Customer = (string.IsNullOrEmpty(data[3]) ? RoznCustomer : data[3]),
                    Order_Sklad = data[1],
                    Order_PodSklad = data[2],
                    Order_Ref = (data[4].Length > 150 ? data[4].Substring(0, 150) : data[4]),
                    Records = new List<DocRecord>()
                };
        }

        private DocRecord AddUpdateDocRecord(Doc doc, string[] data, double Skidka)
        {
            DocRecord rec = doc.Records.FirstOrDefault(x => x.Product_id == data[2]);
            if (rec != null)
            {
                rec.Quantity += Convert.ToDouble(data[4], CultureInfo.InvariantCulture);
                rec.Summa += Convert.ToDouble(data[5], CultureInfo.InvariantCulture);
            }
            else
            {
                rec = new DocRecord()
                {
                    Product_id = data[2],
                    Product_name = data[3],
                    Quantity = Convert.ToDouble(data[4], CultureInfo.InvariantCulture),
                    Summa = Convert.ToDouble(data[5], CultureInfo.InvariantCulture)
                };
                if ((data.Length > 6) && (data[6] == "1"))
                    rec.SkidkaNovTovar = Skidka;
                doc.Records.Add(rec);
            }
            return rec;
        }

        private string[] ParseCVSline(string line)
        {
            List<string> result = new List<string>();
            try
            {
                System.Text.RegularExpressions.Regex pattern = new System.Text.RegularExpressions.Regex(@"
                    # Parse CVS line. Capture next value in named group: 'val'
                    \s*                      # Ignore leading whitespace.
                    (?:                      # Group of value alternatives.
                      ""                     # Either a double quoted string,
                      (?<val>                # Capture contents between quotes.
                        [^""]*(""""[^""]*)*  # Zero or more non-quotes, allowing 
                      )                      # doubled "" quotes within string.
                      ""\s*                  # Ignore whitespace following quote.
                    |  (?<val>[^;]*)         # Or... zero or more non-commas.
                    )                        # End value alternatives group.
                    (?:;)                    # Match end is comma
                    |                        # Or...
                    \s*                      # Ignore leading whitespace.
                    (?:                      # Group of value alternatives.
                      ""                     # Either a double quoted string,
                      (?<val>                # Capture contents between quotes.
                        [^""]*(""""[^""]*)*  # Zero or more non-quotes, allowing 
                      )                      # doubled "" quotes within string.
                      ""\s*                  # Ignore whitespace following quote.
                    |  (?<val>[^;]+)         # Or... one or more non-commas.
                    )                        # End value alternatives group.
                    (?:$)                    # Match end is EOS",
                    System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace);
                System.Text.RegularExpressions.Match matchResult = pattern.Match(line);
                while (matchResult.Success)
                {
                    result.Add(matchResult.Groups["val"].Value.Replace("\"\"", "\""));
                    matchResult = matchResult.NextMatch();
                }
            }
            catch { }
            return result.ToArray();
        }

        [WebMethod(Description = "Создание документов Предварительная Заявка и прочих видов")]
        public string CreateDocuments(byte[] InputFile)
        {
            string FunctionResult = "";
            bool DefaultType = true;
            List<Doc> Docs = new List<Doc>();
            List<string> CustomerList = new List<string>();
            string RoznCustomer = "D0004096";
            List<string> TovarList = new List<string>();
            double SkidkaNovTovar = Convert.ToDouble(ConfigurationManager.AppSettings["SkidkaNovTovar"]);
            if (InputFile.Length > 0)
            {
                using (StreamReader sr = new StreamReader(new MemoryStream(InputFile), System.Text.Encoding.UTF8))
                {
                    string line;
                    Dictionary<string, Doc> OrderNumbers = new Dictionary<string, Doc>();
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] data = ParseCVSline(line);
                        //string[] data = line.Split(';');
                        if (data.Length >= 6)
                        {
                            switch (data[0])
                            {
                                case "0":
                                    Doc doc = NewDoc(data, RoznCustomer);
                                    if (!CustomerList.Contains(doc.Order_Customer))
                                        CustomerList.Add(doc.Order_Customer);
                                    Docs.Add(doc);
                                    break;
                                case "00":
                                    doc = Docs.FirstOrDefault(x => x.Order_Customer == data[3]);
                                    if (doc == null)
                                    {
                                        doc = NewDoc(data, RoznCustomer);
                                        if (!CustomerList.Contains(doc.Order_Customer))
                                            CustomerList.Add(doc.Order_Customer);
                                        Docs.Add(doc);
                                    }
                                    else
                                    {
                                        doc.Order_TotalSum += Convert.ToDouble(data[5], CultureInfo.InvariantCulture);
                                        OrderNumbers.Add(data[1], doc);
                                    }
                                    break;
                                case "1":
                                case "11":
                                    doc = Docs.FirstOrDefault(x => x.Order_id == data[1]);
                                    if (doc != null)
                                    {
                                        DocRecord rec = AddUpdateDocRecord(doc, data, SkidkaNovTovar);
                                        if (!TovarList.Contains(rec.Product_id))
                                            TovarList.Add(rec.Product_id);
                                    }
                                    else
                                    {
                                        doc = OrderNumbers[data[1]];
                                        if (doc != null)
                                        {
                                            DocRecord rec = AddUpdateDocRecord(doc, data, SkidkaNovTovar);
                                            if (!TovarList.Contains(rec.Product_id))
                                                TovarList.Add(rec.Product_id);
                                        }
                                    }
                                    break;
                                case "DH1611&DH2051":
                                    //Реализация + СчФактура
                                    DefaultType = false;
                                    doc = NewDoc(data, RoznCustomer, "1611", "5811");
                                    if (!CustomerList.Contains(doc.Order_Customer))
                                        CustomerList.Add(doc.Order_Customer);
                                    Docs.Add(doc);
                                    Docs.Add(NewDoc(data, RoznCustomer, "2051", "13276", doc));
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
            foreach (Doc doc in Docs)
            {
                if (string.IsNullOrEmpty(doc.Order_type) && (doc.Order_TotalSum != Math.Round(doc.Records.Sum(n => n.Summa), 2)))
                {
                    FunctionResult = AddReplyString(FunctionResult, "Сумма по документу " + doc.Order_id + " не совпадает с суммой всех строк. Загружен не будет.");
                    doc.Order_TotalSum = -1;
                }
            }
            Docs = Docs.Where(x => x.Order_TotalSum >= 0).ToList();

            if (Docs.Count > 0)
            {
                using (SqlConnection connection = new SqlConnection(connDocsStr))
                {
                    string queryString = @"select SC4014.id, SC4014.descr, SC4014.SP4011, SC4014.SP4133, SC131.SP145, (select top 1 value from _1SCONST WITH (NOLOCK) where ID=12006 and OBJID=SC4014.ID and DATE<=convert(varchar(8), @date, 112) order by ID DESC, OBJID DESC, DATE DESC, TIME DESC, DOCID DESC) as postfix from SC4014 (NOLOCK)
                                                inner join SC131 (NOLOCK) on SC131.id = SC4014.SP4011
                                                where SC4014.ISMARK = 0; ";
                    queryString = queryString + @"select SC172.id, SC172.code, SC172.descr, SC172.SP667, ISNULL(sc204.SP1948, '     4   '), ISNULL(sc204.SP1920,'     0   '), SC172.SP9631, SC172.SP12916, ISNULL(SC9633.SP11782,'     0   '), ISNULL(SC426.SP429, 0) from SC172 (NOLOCK)
                                                  left join SC204 (NOLOCK) on sc204.id = sc172.sp667
                                                  left join SC426 (NOLOCK) on SC426.id = sc204.SP1920
                                                  left join SC9633 (NOLOCK) on SC9633.id = sc172.SP9631
                                                  where SC172.code in ({0}); ";
                    queryString = queryString + @"select PARENTEXT, SP11856 from SC11860 (NOLOCK)
                                                  where PARENTEXT in (select id from sc172 (nolock) where code in ({0})) and ISMARK = 0 and (sp11857 = cast('17530101' as datetime) or sp11857 >= cast('" + DateTime.Now.ToString("yyyyMMdd") + "' as datetime)); ";
                    queryString += @"select sc55.id, sc55.code, sc55.descr, IsNull(sc8963.id,''), IsNull(sc8963.code,''), IsNull(sc8963.descr,'') from sc55 (NOLOCK)
                                    left join sc8963 (NOLOCK) on sc8963.parentext = sc55.id
                                    order by sc55.id; ";
                    if (TovarList.Count > 0)
                        queryString = queryString + @"select SC84.id, SC84.code, SC84.descr, SC84.SP94, ISNULL(SC75.SP78, 1), SC84.SP8842 from SC84 (NOLOCK)
                                                  left join SC75 (NOLOCK) on SC75.id = SC84.SP94
                                                  where SC84.code in ({1}); ";
                    string[] paramCustomer = CustomerList.Select(x => "'" + x + "'").ToArray();
                    string[] paramTovar = TovarList.Select(x => "'" + x + "'").ToArray();
                    queryString = string.Format(queryString, string.Join(",", paramCustomer), string.Join(",", paramTovar));

                    SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    command.Parameters.AddWithValue("@date", DateTime.Today);
                    SqlDataReader reader = command.ExecuteReader();
                    List<FirmaDetails> Firms = new List<FirmaDetails>();
                    List<Customer> Customers = new List<Customer>();
                    List<SkladProperties> sklads = new List<SkladProperties>();
                    List<Tovar> Tovars = new List<Tovar>();
                    try
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                FirmaDetails firma = new FirmaDetails();
                                firma.Id = reader.GetString(0);
                                firma.Name = reader.GetString(1);
                                firma.UrLitso = reader.GetString(2);
                                firma.BankSchet = reader.GetString(3);
                                firma.Prefix = reader.GetString(4).Trim();
                                firma.Postfix = reader.GetString(5).Trim();
                                Firms.Add(firma);
                            }
                        }
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                Customer cust = new Customer();
                                cust.Id = reader.GetString(0);
                                cust.Code = reader.GetString(1);
                                cust.Name = reader.GetString(2);
                                cust.Firma = Firms.FirstOrDefault(x => x.Id == Firma.IP_pavlov);
                                cust.MainDogovorId = reader.GetString(3);
                                cust.TypeCostId = reader.GetString(4);
                                cust.SkidkaId = reader.GetString(5);
                                cust.GroupId = reader.GetString(6);
                                cust.Export = Convert.ToBoolean(reader.GetValue(7), CultureInfo.InvariantCulture);
                                cust.KolonkaSkidki = reader.GetString(8);
                                cust.SkidkaValue = Convert.ToDouble(reader.GetValue(9), CultureInfo.InvariantCulture);
                                Customers.Add(cust);
                            }
                        }
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                Customer cust = Customers.FirstOrDefault(x => x.Id == reader.GetString(0));
                                if ((cust != null) && (reader.GetString(1) == Firma.StinPlus))
                                {
                                    cust.Firma = Firms.FirstOrDefault(x => x.Id == Firma.StinPlus);
                                }
                            }
                        }
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                SkladProperties sklad = sklads.FirstOrDefault(x => x.ID == reader.GetString(0));
                                if (sklad != null)
                                    sklad.PodSklads.Add(new SkladProperties
                                    {
                                        ID = reader.GetString(3),
                                        Code = reader.GetString(4),
                                        Name = reader.GetString(5)
                                    });
                                else
                                    sklads.Add(new SkladProperties
                                    {
                                        ID = reader.GetString(0),
                                        Code = reader.GetString(1),
                                        Name = reader.GetString(2),
                                        PodSklads = new List<SkladProperties> { new SkladProperties {
                                            ID = reader.GetString(3),
                                            Code = reader.GetString(4),
                                            Name = reader.GetString(5)
                                        } }
                                    });
                            }
                        }
                        if (reader.NextResult())
                        {
                            while (reader.Read())
                            {
                                Tovar tov = new Tovar();
                                tov.Id = reader.GetString(0);
                                tov.Code = reader.GetString(1);
                                tov.Name = reader.GetString(2);
                                tov.EdId = reader.GetString(3);
                                tov.koff = Convert.ToDouble(reader.GetValue(4), CultureInfo.InvariantCulture);
                                tov.ProizId = reader.GetString(5);
                                Tovars.Add(tov);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        FunctionResult = AddReplyString(FunctionResult, e.Message);
                    }
                    finally
                    {
                        reader.Close();
                    }
                    if ((Tovars.Count > 0) && (Customers.Count > 0))
                        CalculateCostHelper(Tovars, Customers);
                    try
                    {
                        List<string> DbSigns = new List<string>();
                        using (SqlCommand commandSigns = new SqlCommand("select dbsign from _1sdbset", connection))
                        {
                            using (SqlDataReader readerSigns = commandSigns.ExecuteReader())
                            {
                                while (readerSigns.Read())
                                {
                                    DbSigns.Add(readerSigns["dbsign"].ToString().Trim());
                                }
                            }
                        }

                        string PrefixDb = "E";
                        string queryPrefix = @"select top 1 value from _1sconst with (NOLOCK) 
                                                    where id = 3701 and objid = '     0   ' and date <= '1753-01-01 00:00:00'
                                                    order by id desc, objid desc, date desc, time desc, docid desc; ";
                        using (SqlCommand commandPrefix = new SqlCommand(queryPrefix, connection))
                        {
                            using (SqlDataReader readerPrefix = commandPrefix.ExecuteReader())
                            {
                                while (readerPrefix.Read())
                                {
                                    PrefixDb = readerPrefix["value"].ToString().Trim();
                                }
                            }
                        }
                        if (DefaultType)
                        {
                            //queryString = @"SET IDENTITY_INSERT _1sjourn ON
                            queryString = @"
                                    insert into _1sjourn (IDJOURNAL,IDDOC,IDDOCDEF,APPCODE,DATE_TIME_IDDOC,DNPREFIX,DOCNO,CLOSED,ISMARK,ACTCNT,VERSTAMP,
                                    RF639,RF464,RF4667,RF4674,RF635,RF3549,RF4343,RF8677,RF8696,RF405,RF328,RF351,RF2964,RF4335,RF4314,RF2351,RF438,RF4480,RF8888,RF8894,
                                    RF9143,RF9469,RF9531,RF9596,RF9972,RF9981,RF9989,RF10305,RF10313,RF10318,RF10324,RF10471,RF10476,RF11049,RF11055,RF11495,RF11973,RF12351,
                                    RF12406,RF12413,RF12503,RF12566,RF12618,RF12791,RF12815,RF13979,RF14021,
                                    SP74,SP798,SP4056,SP5365,SP8662,SP8663,SP8664,SP8665,SP8666,SP8720,SP8723,
                                    ds1946,ds4757,ds5722) 
                                    values(4588,@iddoc,@iddocdef,1,@date_time_iddoc,@DNPREFIX,@number,0,0,0,0,
                                    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                                    '    20D  ','     0   ',@firma,@urlitso,'D','D;Экран',@SP8664,'D;ПредварительнаяЗаявка',@firmaName,'','',       
                                    0,0,0)
                                    insert into DH12747 (IDDOC,SP12711,SP12712,SP12713,SP12714,SP12715,SP12716,SP12717,SP12718,SP12719,SP12720,SP12721,SP12722,SP12723,
                                    SP12724,SP12725,SP12726,SP12727,SP12728,SP12729,SP12730,SP12731,SP12732,SP12733,SP12734,SP12735,SP14007,SP12741,SP12743,SP12745,SP660)
                                    values (@iddoc,'   0     0   ',@bankSchet,@cust_id,@dogovor,'     1   ',1,1,1,0,1,@SumDoc,@CostType,@skidka,@dateOtg,@dateOtg,
                                    @sklad,'     0   ','     0   ',1,0,'   80N   ','',0,'','','     0   ',0,0,0,@ref)";
                                    //SET IDENTITY_INSERT _1sjourn OFF ";

                            string queryDT = @"insert into DT12747 (IDDOC,LINENO_,SP12736,SP12737,SP12738,SP12739,SP12740,SP12741,SP12742,SP12743,SP12744,SP12745,SP13041)
                                                       values (@iddoc,@lineno,@tovar,@kolvo,@ed,@koff,@cost,@sum,@nds,@sum_nds,'     1   ',0,@skidkanovtovar);";
                            string queryCheckNumber = @"select closed from _1sjourn where iddocdef = @iddocdef and docno = @number and left(date_time_iddoc,4) = @year";

                            using (SqlCommand commandInsert = new SqlCommand(queryString, connection))
                            {
                                commandInsert.Parameters.Add("@iddoc", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@iddocdef", SqlDbType.Int);
                                commandInsert.Parameters.Add("@date_time_iddoc", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@DNPREFIX", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@number", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@firma", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@urlitso", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@SP8664", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@firmaName", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@bankSchet", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@cust_id", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@dogovor", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@SumDoc", SqlDbType.Float);
                                commandInsert.Parameters.Add("@CostType", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@skidka", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@dateOtg", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@sklad", SqlDbType.VarChar);
                                commandInsert.Parameters.Add("@ref", SqlDbType.VarChar);
                                using (SqlCommand commandCheck = new SqlCommand(queryCheckNumber, connection))
                                {
                                    commandCheck.Parameters.Add("@iddocdef", SqlDbType.Int);
                                    commandCheck.Parameters.Add("@number", SqlDbType.VarChar);
                                    commandCheck.Parameters.Add("@year", SqlDbType.VarChar);

                                    int iddocdef = 12747;
                                    foreach (Doc doc in Docs)
                                    {
                                        Customer cust = Customers.FirstOrDefault(x => x.Code == doc.Order_Customer);
                                        if (cust != null)
                                        {
                                            FirmaDetails firma = cust.Firma;
                                            if (doc.Order_Firma == Firma.Stin_service)
                                                firma = Firms.FirstOrDefault(x => x.Id == Firma.Stin_service);
                                            string DocNumber = "E" + firma.Prefix + doc.Order_id.PadLeft(8, '0');
                                            commandCheck.Parameters["@iddocdef"].Value = iddocdef;
                                            commandCheck.Parameters["@number"].Value = DocNumber;
                                            commandCheck.Parameters["@year"].Value = doc.Order_date.ToString("yyyy");
                                            bool CanCreateDoc = false;
                                            using (SqlDataReader checker = commandCheck.ExecuteReader())
                                            {
                                                CanCreateDoc = !checker.HasRows;
                                            }
                                            if (CanCreateDoc)
                                            {
                                                string num36 = GenerateIdDoc(connection);
                                                if (!string.IsNullOrEmpty(num36))
                                                {
                                                    commandInsert.Parameters["@iddoc"].Value = num36;
                                                    commandInsert.Parameters["@iddocdef"].Value = iddocdef;
                                                    commandInsert.Parameters["@date_time_iddoc"].Value = doc.Order_date.ToString("yyyyMMdd") + Encode36(180000000) + num36;
                                                    commandInsert.Parameters["@DNPREFIX"].Value = iddocdef.ToString().PadLeft(10) + doc.Order_date.ToString("yyyy").PadRight(8);
                                                    commandInsert.Parameters["@number"].Value = DocNumber;
                                                    commandInsert.Parameters["@firma"].Value = firma.Id;
                                                    commandInsert.Parameters["@urlitso"].Value = firma.UrLitso;
                                                    commandInsert.Parameters["@SP8664"].Value = "D;" + cust.Name.Substring(0, 28);
                                                    commandInsert.Parameters["@firmaName"].Value = "D;" + (firma.Name.Length > 28 ? firma.Name.Substring(0, 28) : firma.Name);
                                                    commandInsert.Parameters["@bankSchet"].Value = firma.BankSchet;
                                                    commandInsert.Parameters["@cust_id"].Value = cust.Id;
                                                    commandInsert.Parameters["@dogovor"].Value = cust.MainDogovorId;
                                                    commandInsert.Parameters["@CostType"].Value = cust.TypeCostId;
                                                    commandInsert.Parameters["@skidka"].Value = cust.SkidkaId;
                                                    commandInsert.Parameters["@dateOtg"].Value = DateTime.Now.ToString("yyyyMMdd");
                                                    commandInsert.Parameters["@sklad"].Value = Sklad.Ekran;
                                                    commandInsert.Parameters["@ref"].Value = doc.Order_Ref;

                                                    double SumDoc = 0;
                                                    using (SqlCommand commandDT = new SqlCommand(queryDT, connection))
                                                    {
                                                        commandDT.Parameters.AddWithValue("@iddoc", num36);
                                                        commandDT.Parameters.Add("@lineno", SqlDbType.Int);
                                                        commandDT.Parameters.Add("@tovar", SqlDbType.VarChar);
                                                        commandDT.Parameters.Add("@kolvo", SqlDbType.Float);
                                                        commandDT.Parameters.Add("@ed", SqlDbType.VarChar);
                                                        commandDT.Parameters.Add("@koff", SqlDbType.Float);
                                                        commandDT.Parameters.Add("@cost", SqlDbType.Float);
                                                        commandDT.Parameters.Add("@sum", SqlDbType.Float);
                                                        commandDT.Parameters.Add("@nds", SqlDbType.VarChar);
                                                        commandDT.Parameters.Add("@sum_nds", SqlDbType.Float);
                                                        commandDT.Parameters.Add("@skidkanovtovar", SqlDbType.Float);
                                                        int lineno = 0;
                                                        foreach (DocRecord rec in doc.Records)
                                                        {
                                                            Tovar tov = Tovars.FirstOrDefault(x => x.Code == rec.Product_id);
                                                            if (tov != null)
                                                            {
                                                                lineno++;
                                                                double summa = CalculateSumma(tov, cust, rec.Quantity, rec.SkidkaNovTovar, (doc.Order_Firma == Firma.Stin_service));
                                                                SumDoc += summa;
                                                                commandDT.Parameters["@lineno"].Value = lineno;
                                                                commandDT.Parameters["@tovar"].Value = tov.Id;
                                                                commandDT.Parameters["@kolvo"].Value = rec.Quantity;
                                                                commandDT.Parameters["@ed"].Value = tov.EdId;
                                                                commandDT.Parameters["@koff"].Value = tov.koff;
                                                                commandDT.Parameters["@cost"].Value = (rec.Quantity > 0 ? summa / rec.Quantity : summa);
                                                                commandDT.Parameters["@sum"].Value = summa;
                                                                if (cust.Export)
                                                                {
                                                                    commandDT.Parameters["@nds"].Value = HDS.HDS_0;
                                                                    commandDT.Parameters["@sum_nds"].Value = 0;
                                                                }
                                                                else
                                                                {
                                                                    if (DateTime.Today.Year < 2019)
                                                                    {
                                                                        commandDT.Parameters["@nds"].Value = HDS.HDS_18;
                                                                        commandDT.Parameters["@sum_nds"].Value = summa * 18 / 118;
                                                                    }
                                                                    else
                                                                    {
                                                                        commandDT.Parameters["@nds"].Value = HDS.HDS_20;
                                                                        commandDT.Parameters["@sum_nds"].Value = summa * 20 / 120;
                                                                    }
                                                                }
                                                                commandDT.Parameters["@skidkanovtovar"].Value = rec.SkidkaNovTovar;

                                                                commandDT.ExecuteNonQuery();

                                                            }
                                                            else
                                                                FunctionResult = AddReplyString(FunctionResult, "Номенклатуры " + rec.Product_name + " код (" + rec.Product_id + ") не обнаружено. В документе " + DocNumber + " от " + doc.Order_date.ToString("dd.MM.yyyy") + " эта строка пропущена.");
                                                        }
                                                    }

                                                    commandInsert.Parameters["@SumDoc"].Value = SumDoc;
                                                    commandInsert.ExecuteNonQuery();

                                                    using (SqlCommand commandUpdate = new SqlCommand("_1sp_DH12747_UpdateTotals", connection)
                                                    {
                                                        CommandType = CommandType.StoredProcedure
                                                    })
                                                    {
                                                        commandUpdate.Parameters.AddWithValue("@iddoc", num36);
                                                        commandUpdate.ExecuteNonQuery();
                                                        FunctionResult = AddReplyString(FunctionResult, "Создан документ ПредварительнаяЗаявка " + DocNumber + " от " + doc.Order_date.ToString("dd.MM.yyyy"));
                                                    }
                                                }
                                                else
                                                    FunctionResult = AddReplyString(FunctionResult, "Ошибка получения IdDoc. Документ " + DocNumber + " от " + doc.Order_date.ToString("dd.MM.yyyy") + " загружен не будет.");
                                            }
                                            else
                                                FunctionResult = AddReplyString(FunctionResult, "Документ " + DocNumber + " от " + doc.Order_date.ToString("yyyy") + " уже существует в БД и загружен не будет.");
                                        }
                                        else
                                            FunctionResult = AddReplyString(FunctionResult, "Контрагента с кодом " + doc.Order_Customer + " не обнаружено. Заявка " + doc.Order_id + " загружена не будет.");
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (Doc doc in Docs)
                            {
                                string Prefix = PrefixDb + "W";
                                string Postfix = "";
                                FirmaDetails firma = Firms.FirstOrDefault(x => x.Id == Firma.IP_pavlov);
                                if (doc.Order_Firma == Firma.Stin_service)
                                    firma = Firms.FirstOrDefault(x => x.Id == Firma.Stin_service);
                                Customer cust = Customers.FirstOrDefault(x => x.Code == doc.Order_Customer);
                                if (cust != null)
                                {
                                    if (doc.Order_Firma != Firma.Stin_service)
                                        firma = cust.Firma;
                                    Prefix = PrefixDb + firma.Prefix;
                                    Postfix = firma.Postfix;
                                    if (Postfix == "@")
                                        Postfix = "1";
                                }
                                SkladProperties sklad = sklads.FirstOrDefault(x => x.ID == Sklad.Ekran);
                                SkladProperties podSklad = null;
                                if (doc.Order_Sklad != null)
                                    sklad = sklads.FirstOrDefault(x => x.Code == doc.Order_Sklad);
                                if ((doc.Order_PodSklad != null) && (sklad != null))
                                    podSklad = sklad.PodSklads.FirstOrDefault(x => x.Code == doc.Order_PodSklad);

                                string dnPrefix = doc.Order_numerator.PadLeft(10) + DateTime.Today.ToString("yyyy").PadRight(8);
                                string queryDocNo = "select top 1 DOCNO from _1SJOURN WITH (NOLOCK) where DNPREFIX='" + dnPrefix + "' and DOCNO>='" + Prefix + "' and substring(DOCNO,1," + Prefix.Length.ToString() + ")='" + Prefix + "' order by DNPREFIX DESC, DOCNO DESC";
                                using (SqlTransaction tran = connection.BeginTransaction())
                                {
                                    string NewDocNo = Prefix + "00000001";
                                    if ((doc.Order_type == "2051") && (!string.IsNullOrEmpty(Postfix)))
                                        NewDocNo = Prefix + "000001/" + Postfix;
                                    using (SqlCommand commandDocNo = new SqlCommand(queryDocNo, connection, tran))
                                    {
                                        using (SqlDataReader readerDocNo = commandDocNo.ExecuteReader())
                                        {
                                            if (readerDocNo.Read())
                                            {
                                                string postfix = "";
                                                NewDocNo = readerDocNo["DOCNO"].ToString().Trim().Substring(Prefix.Length);
                                                if (doc.Order_type == "2051")
                                                {
                                                    postfix = "/" + Postfix;
                                                    NewDocNo = NewDocNo.Substring(0, NewDocNo.Length - postfix.Length);
                                                }
                                                NewDocNo = Prefix + (postfix == "" ? (Convert.ToInt32(NewDocNo) + 1).ToString().PadLeft(8, '0') : (Convert.ToInt32(NewDocNo) + 1).ToString().PadLeft(6, '0')) + postfix;
                                            }
                                        }
                                    }
                                    string num36 = GenerateIdDoc(connection, tran);
                                    if (!string.IsNullOrEmpty(num36))
                                    {
                                        doc.Order_id = num36;
                                        queryString = @" 
                                            insert into _1sjourn (IDJOURNAL,IDDOC,IDDOCDEF,APPCODE,DATE_TIME_IDDOC,DNPREFIX,DOCNO,CLOSED,ISMARK,ACTCNT,VERSTAMP,
                                            RF639,RF464,RF4667,RF4674,RF635,RF3549,RF4343,RF8677,RF8696,RF405,RF328,RF351,RF2964,RF4335,RF4314,RF2351,RF438,RF4480,RF8888,RF8894,
                                            RF9143,RF9469,RF9531,RF9596,RF9972,RF9981,RF9989,RF10305,RF10313,RF10318,RF10324,RF10471,RF10476,RF11049,RF11055,RF11495,RF11973,RF12351,
                                            RF12406,RF12413,RF12503,RF12566,RF12618,RF12791,RF12815,RF13979,RF14021,
                                            SP74,SP798,SP4056,SP5365,SP8662,SP8663,SP8664,SP8665,SP8666,SP8720,SP8723,
                                            ds1946,ds4757,ds5722) 
                                            values(@idjourn,@iddoc,@iddocdef,1,@date_time_iddoc,@DNPREFIX,@number,0,0,0,0,
                                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                                            '    20D  ','     0   ',@firma,@urlitso,@prefixDB,@prefixSklad,@SP8664,@prefixVid,@firmaName,'','',       
                                            0,0,0)";
                                        if (doc.Order_type == "1611")
                                        {
                                            queryString += @"insert into DH1611 (IDDOC,SP3338,SP1587,SP1593,SP1583,SP1584,SP1585,SP1586,SP8818,SP1589,SP1590,SP1591,SP1592,SP1595,SP1596,SP1594,SP1588,SP7550,SP8673,SP8691,SP9000,SP9001,SP9537,SP9538,SP10268,SP11568,SP11569,SP1604,SP1605,SP1606,SP660)
                                            values(@iddoc,@codeOper,@docOsn,@sklad_id,@cust_id,@dogovor,'     1   ',1.0000,0,1,1,0,1,@CostType,@skidka,0,@dateOtg,0,'     0   ',0,@podSklad,0,0,1,'     0   ','','',0,0,0,@ref) ";
                                        }
                                        else
                                        {
                                            queryString += @"insert into DH2051 (IDDOC,SP2034,SP2035,SP2038,SP4394,SP4395,SP4396,SP4397,SP4398,SP4399,SP4400,SP4401,SP2036,SP2037,SP7847,SP2043,SP2044,SP2045,SP660)
                                            values(@iddoc,@cust_id,@dogovor,@docOsn,1,0,1,1,0,1,'','1753-01-01 00:00:00','     1   ',1.0000,0,0,0,0,@ref) ";
                                        }

                                        using (SqlCommand commandInsert = new SqlCommand(queryString, connection, tran))
                                        {
                                            commandInsert.Parameters.AddWithValue("@idjourn", (doc.Order_type == "1611" ? 4588 : 4601));
                                            commandInsert.Parameters.AddWithValue("@iddoc", num36);
                                            commandInsert.Parameters.AddWithValue("@iddocdef", Convert.ToInt32(doc.Order_type));
                                            commandInsert.Parameters.AddWithValue("@date_time_iddoc", DateTime.Today.ToString("yyyyMMdd") + Encode36(648000000) + num36);
                                            commandInsert.Parameters.AddWithValue("@DNPREFIX", dnPrefix);
                                            commandInsert.Parameters.AddWithValue("@number", NewDocNo);
                                            commandInsert.Parameters.AddWithValue("@firma", firma.Id);
                                            commandInsert.Parameters.AddWithValue("@urlitso", firma.UrLitso);
                                            commandInsert.Parameters.AddWithValue("@prefixDB", PrefixDb);
                                            commandInsert.Parameters.AddWithValue("@prefixSklad", PrefixDb + ";" + (sklad.Name.Length > 28 ? sklad.Name.Substring(0, 28) : sklad.Name));
                                            commandInsert.Parameters.AddWithValue("@SP8664", PrefixDb + ";" + cust.Name.Substring(0, 28));
                                            commandInsert.Parameters.AddWithValue("@prefixVid", PrefixDb + ";" + (doc.Order_type == "1611" ? "Реализация" : "СчетФактураВыданный"));
                                            commandInsert.Parameters.AddWithValue("@firmaName", PrefixDb + ";" + (firma.Name.Length > 28 ? firma.Name.Substring(0, 28) : firma.Name));
                                            commandInsert.Parameters.AddWithValue("@codeOper", "   16S   ");
                                            commandInsert.Parameters.AddWithValue("@sklad_id", sklad.ID);
                                            commandInsert.Parameters.AddWithValue("@cust_id", cust.Id);
                                            commandInsert.Parameters.AddWithValue("@dogovor", cust.MainDogovorId);
                                            commandInsert.Parameters.AddWithValue("@CostType", cust.TypeCostId);
                                            commandInsert.Parameters.AddWithValue("@skidka", cust.SkidkaId);
                                            commandInsert.Parameters.AddWithValue("@dateOtg", DateTime.Now.ToString("yyyyMMdd"));
                                            if (podSklad != null)
                                                commandInsert.Parameters.AddWithValue("@podSklad", podSklad.ID);
                                            else
                                                commandInsert.Parameters.AddWithValue("@podSklad", "     0   ");
                                            commandInsert.Parameters.AddWithValue("@ref", doc.Order_Ref);
                                            if ((doc.Order_parent != null) && (!string.IsNullOrEmpty(doc.Order_parent.Order_id)))
                                                commandInsert.Parameters.AddWithValue("@docOsn", Encode36(Convert.ToInt64(doc.Order_parent.Order_type)).PadLeft(4) + doc.Order_parent.Order_id);
                                            else
                                                commandInsert.Parameters.AddWithValue("@docOsn", "   0     0   ");
                                            commandInsert.ExecuteNonQuery();
                                        }
                                        if (doc.Order_type == "2051")
                                        {
                                            Fill_1SCRDOC(
                                                connection,
                                                tran,
                                                0, //подчиненные документы
                                                "O1" + Encode36(Convert.ToInt64(doc.Order_parent.Order_type)).PadLeft(4) + doc.Order_parent.Order_id,
                                                DateTime.Today.ToString("yyyyMMdd") + Encode36(648000000) + num36,
                                                num36
                                                );
                                        }
                                        if (doc.Order_type == "1611")
                                        {
                                            Fill_1SCRDOC(
                                                connection,
                                                tran,
                                                4747, // графа отбора Склад
                                                "B1" + Encode36(55).PadLeft(4) + sklad.ID,
                                                DateTime.Today.ToString("yyyyMMdd") + Encode36(648000000) + num36,
                                                num36
                                                );
                                        }
                                        Fill_1SCRDOC(
                                            connection,
                                            tran,
                                            862, // графа отбора Контрагент
                                            "B1" + Encode36(172).PadLeft(4) + cust.Id,
                                            DateTime.Today.ToString("yyyyMMdd") + Encode36(648000000) + num36,
                                            num36
                                            );
                                        foreach (string dbSign in DbSigns)
                                        {
                                            if (dbSign != PrefixDb)
                                                RegisterUpdate(
                                                    connection,
                                                    tran,
                                                    dbSign.PadRight(3),
                                                    Convert.ToInt32(doc.Order_type),
                                                    num36,
                                                    " "
                                                    );
                                        }
                                        FunctionResult = AddReplyString(FunctionResult, "Создан документ " + (doc.Order_type == "1611" ? "Реализация " : "СчетФактураВыданный ") + NewDocNo + " от " + DateTime.Today.ToString("dd.MM.yyyy"));
                                    }
                                    tran.Commit();
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        FunctionResult = AddReplyString(FunctionResult, e.Message);
                    }
                }
            }
            else
            {
                FunctionResult = AddReplyString(FunctionResult, "Не обнаружено данных для загрузки.");
            }

            return FunctionResult;
        }

        [WebMethod(Description = "Получение отчета по всем купленным товарам клиентами для Web в виде CSV файла.")]
        public PriceList GetFullTovarProdagiForWeb()
        {
            PriceList reply = new PriceList();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            using (SqlConnection connection = new SqlConnection(connWebStr))
            {
                string queryString = @"SELECT SC172.code FROM SC172 (NOLOCK)
                                        inner JOIN _1SCONST As TabConst (NOLOCK)
                                        ON ((SC172.ID = TabConst.OBJID)
                                            AND (TabConst.ID = 12329)
	                                        AND (TabConst.VALUE is not null)
	                                        AND (TabConst.VALUE = 1)
                                            AND (TabConst.DATE =
                                                (SELECT MAX(TabConstl.DATE)
                                                FROM _1SCONST AS TabConstl (NOLOCK)
                                                WHERE TabConstl.OBJID = TabConst.OBJID
                                                    AND TabConstl.ID = TabConst.ID
                                                )
                                            )
                                        )";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 120;
                connection.Open();
                string Line;
                SqlDataReader reader = command.ExecuteReader();
                string CustomersCodes = "";
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            if (string.IsNullOrEmpty(CustomersCodes))
                                CustomersCodes += reader.GetString(0).Trim();
                            else
                                CustomersCodes += "," + reader.GetString(0).Trim();
                        }
                    }
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
                if (!string.IsNullOrEmpty(CustomersCodes))
                {
                    using (MemoryStream stream = new MemoryStream())
                    {
                        using (TextWriter tw = new StreamWriter(stream, System.Text.Encoding.UTF8))
                        {
                            NomenklaturaProdanaReply npr = NomenklaturaProdana(CustomersCodes);
                            if (npr.ReplyState.State == "Success")
                            {
                                foreach (CustTovar t in npr.Tovars)
                                {
                                    Line = t.CustCode + ";";
                                    Line = Line + t.Code + ";";
                                    Line = Line + CheckForQuotes(t.Name);
                                    tw.WriteLine(Line);
                                }
                            }
                        }
                        reply.ExcelFile = stream.ToArray();
                    }
                    reply.ReplyState.State = "Success";
                    reply.ReplyState.Description = "Информация успешно извлечена из БД";
                }
            }

            return reply;

        }

        [WebMethod(Description = "Получение отчета по всем товарам для Web в виде CSV файла. Товары, не попавшие в файл считаются удаленными. Формирование ответа занимает 5-6 минут")]
        public PriceList GetFullTovarForWeb(bool IncludeProperties = true)
        {
            PriceList reply = new PriceList();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            using (SqlConnection connection = new SqlConnection(connWebStr))
            {
                string querySystem = "select top 1 curdate from _1ssystem;";
                string queryString = @"select 
                                            harak.parentext, types.descr, harak.SP12315
                                        from SC12317 as harak (NOLOCK)
                                        inner join SC12312 as types (NOLOCK) on types.id = harak.SP12314
                                        where harak.ismark = 0 
                                        ;";
                queryString = queryString + @"select nomenk.id,
                                        nomenk.code, nomenk.descr, nomenk.SP8848,
                                        proiz.descr, proiz_gr.descr, nomenk.SP85, ISNULL(oksm.descr, ''),
                                        IsNull(Gr1.Descr, ''), IsNull(Gr2.Descr, ''), IsNull(Gr3.Descr, ''),
                                        IsNull(R.ostatok, 0), IsNull(TabConstWeb.Value, 0), 
                                        IsNull(TabConstOpt.Value, 0), IsNull(TabConst4.Value, 0), IsNull(TabConstRoz.Value, 0), IsNull(TabConst6.Value, 0),
                                        nomenk.SP10397, IsNull(RR.rezerv, 0) as rezerv, IsNull(nomenk.SP10406, 1) as kolUpak, nomenk.SP208, IsNull(R_OstPrih.OstatokPrih, 0) as ostPrih,
                                        IsNull(R_DS.ostatok,0) as OstatokDS, IsNull(RR_DS.rezerv,0) as RezervDS,
                                        nomenk.SP5066 as TolkoRozn, nomenk.SP13277 as YandexMarket, IsNull(Rav.avSp, 0) as avSp
                                        from sc84 as nomenk (NOLOCK)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = nomenk.SP8842
                                        left join SC12934 as proiz_gr (NOLOCK) on proiz_gr.id = proiz.SP12932 
                                        LEFT JOIN _1SCONST TabConst (NOLOCK) ON nomenk.ID = TabConst.OBJID AND TabConst.ID = 5012 AND 
                                            TabConst.DATE = (SELECT MAX(TabConstl.DATE) FROM _1SCONST AS TabConstl (NOLOCK)
                                                                WHERE TabConstl.OBJID = TabConst.OBJID AND TabConstl.ID = TabConst.ID)
                                        left join SC566 as oksm (NOLOCK) on oksm.id = TabConst.value
                                        LEFT JOIN _1SCONST TabConstWeb (NOLOCK) ON nomenk.ID = TabConstWeb.OBJID AND TabConstWeb.ID = 12305 AND 
                                            TabConstWeb.DATE = (SELECT MAX(TabConstWebl.DATE) FROM _1SCONST AS TabConstWebl (NOLOCK)
                                                                WHERE TabConstWebl.OBJID = TabConstWeb.OBJID AND TabConstWebl.ID = TabConstWeb.ID)
                                        LEFT OUTER JOIN sc84 As Gr3 (NOLOCK) ON nomenk.PARENTID = Gr3.ID
                                        LEFT OUTER JOIN sc84 As Gr2 (NOLOCK) ON Gr3.PARENTID = Gr2.ID
                                        LEFT OUTER JOIN sc84 As Gr1 (NOLOCK) ON Gr2.PARENTID = Gr1.ID 
                                        left join (
                                                    select sp408, Sum(sp411) as ostatok from rg405 (NOLOCK)
                                                        where period = @Period and sp418 = @Ekran and sp4062 in ({1})
                                                    group by sp408) as R on R.sp408 = nomenk.id
                                        left join (
                                                    select SP4477, Sum(SP4479) as rezerv from RG4480 (NOLOCK)
                                                        where period = @Period and SP4476 = @Ekran and sp4475 in ({1})
                                                    group by SP4477) as RR on RR.SP4477 = nomenk.id
                                        left join (
                                                    select SP11050, Sum(SP11054) as avSp from RG11055 (NOLOCK)
                                                        where period = @Period and SP11051 = @Ekran
                                                    group by SP11050) as Rav on Rav.SP11050 = nomenk.id
                                        left join (
                                                    select sp408, Sum(sp411) as ostatok from rg405 (NOLOCK)
                                                        where period = @Period and sp418 = @DomSad and sp4062 in ({1})
                                                    group by sp408) as R_DS on R_DS.sp408 = nomenk.id
                                        left join (
                                                    select SP4477, Sum(SP4479) as rezerv from RG4480 (NOLOCK)
                                                        where period = @Period and SP4476 = @DomSad and sp4475 in ({1})
                                                    group by SP4477) as RR_DS on RR_DS.SP4477 = nomenk.id
                                        left join (
                                                    select SP466, Sum(SP4471) as OstatokPrih from RG464 (NOLOCK)
                                                        where period = @Period and SP4467 in ({1})
                                                    group by SP466) as R_OstPrih on R_OstPrih.SP466 = nomenk.id
                                        left join SC319 as SpCost (NOLOCK) on ((nomenk.ID = SpCost.PARENTEXT) AND SpCost.SP327 = @CostType and (SpCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstOpt (NOLOCK) ON SpCost.ID = TabConstOpt.OBJID AND TabConstOpt.ID = 324 AND 
                                            TabConstOpt.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstOpt.OBJID and id = TabConstOpt.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCostRoz (NOLOCK) on ((nomenk.ID = SpCostRoz.PARENTEXT) AND SpCostRoz.SP327 = @CostRoznType and (SpCostRoz.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstRoz (NOLOCK) ON SpCostRoz.ID = TabConstRoz.OBJID AND TabConstRoz.ID = 324 AND 
                                            TabConstRoz.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstRoz.OBJID and id = TabConstRoz.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCostSP (NOLOCK) on ((nomenk.ID = SpCostSP.PARENTEXT) and (SpCostSP.SP327 = @CostSPType) and (SpCostSP.IsMark = 0))
                                        left join SC8904 as SpSklCost (NOLOCK) on ((SpCostSP.ID = SpSklCost.PARENTEXT) and (SpSklCost.SP8901 = @Ekran) and (SpSklCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst4 (NOLOCK) ON SpSklCost.ID = TabConst4.OBJID AND TabConst4.ID = 8902 AND 
                                            TabConst4.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst4.OBJID and id = TabConst4.id
                                                                  order by date desc, time desc)
                                        left join SC8904 as SpSklCostR (NOLOCK) on ((SpCostSP.ID = SpSklCostR.PARENTEXT) and (SpSklCostR.SP8901 = @SklR) and (SpSklCostR.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst6 (NOLOCK) ON SpSklCostR.ID = TabConst6.OBJID AND TabConst6.ID = 8902 AND 
                                            TabConst6.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst6.OBJID and id = TabConst6.id
                                                                  order by date desc, time desc)
                                        where nomenk.ismark = 0 and nomenk.isfolder = 2 and " + SprInclude("sc84", "nomenk", "{0}", false) +
                                           "and right(RTRIM(proiz.descr),1) <> '*' " +
                                           //"and right(RTRIM(nomenk.descr),3) <> ' БУ' and right(RTRIM(nomenk.descr),7) <> ' УЦЕНКА' " +
                                           "order by nomenk.descr";

                List<string> BadFolders = new List<string>();
                BadFolders.Add(Nomenklatura.Musor); //Мусор код S00012726
                BadFolders.Add(Nomenklatura.Uchenka); //Уцененные товары код D00000015
                BadFolders.Add(Nomenklatura.ForMaster); //Для мастерской код A00003191
                BadFolders.Add(Nomenklatura.PryamProd); //Прямые продажи код D00027231
                BadFolders.Add(Nomenklatura.TorgOborud); //Торговое оборудование и рекламные материалы код D00007946
                string[] paramBF = BadFolders.Select(x => "'" + x + "'").ToArray();
                List<string> FirmaList = new List<string> { Firma.IP_pavlov, Firma.StinPlus, Firma.Stin_service };
                string[] paramFirma = FirmaList.Select(x => "'" + x + "'").ToArray();
                queryString = string.Format(queryString, string.Join(",", paramBF), string.Join(",", paramFirma));

                SqlCommand commandSystem = new SqlCommand(querySystem, connection);
                connection.Open();

                SqlDataReader reader = commandSystem.ExecuteReader();
                DateTime dateTA = DateTime.Now;
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                            dateTA = reader.GetDateTime(0);
                    }
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }

                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 2000;
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                DateTime startDate = new DateTime(dateTA.Year, dateTA.Month, 1);
                command.Parameters["@Period"].Value = startDate;
                command.Parameters.AddWithValue("@Ekran", Sklad.Ekran); //оптовый склад Экран
                command.Parameters.AddWithValue("@DomSad", Sklad.DomSad); //склад "Дом и Сад" (Заводское)
                command.Parameters.AddWithValue("@CostType", CostType.Opt);  //тип цен Оптовая
                command.Parameters.AddWithValue("@CostRoznType", CostType.Rozn);  //тип цен Розничная
                command.Parameters.AddWithValue("@CostSPType", CostType.SP); //тип цен Сп
                command.Parameters.AddWithValue("@SklR", Sklad.Gastello); //розничный склад Гастелло-Инструмент
                //command.Parameters.AddWithValue("@Firma", Firma.IP_pavlov); //фирма ИП Павлов

                //command.Parameters.AddWithValue("@Code1", "K00035589");
                //command.Parameters.AddWithValue("@Code2", "K00035554");
                //command.Parameters.AddWithValue("@Code3", "P00003089");
                //command.Parameters.AddWithValue("@Code4", "D00024119");
                //command.Parameters.AddWithValue("@Code5", "D00024079");

                string Header = "Код;Название;Краткое описание;Производитель;Артикул;";
                if (IncludeProperties)
                    Header = Header + "Мощность, Вт.;Мощность, кВт.;Мощность, л.с.;Напряжение, В.;Напряжение аккумулятора, В.;Емкость аккумулятора, А.ч.;Тип аккумулятора;Гарантия в месяцах;Вес нетто, кг.;";
                Header = Header + "Страна происхождения;Раздел 1;Раздел 2; Раздел 3; Остаток;";
                Header = Header + "Розничная;Розничная спец;Оптовая;Оптовая спец;Валюта 1;Валюта 2;Валюта 3;Валюта 4;";
                Header = Header + "Мин розничная цена;Мин оптовая цена;Размер скидки розн;Размер скидки опт;Количество для заказа (опт);";
                Header = Header + "Яндекс.Маркет;Только для розницы;";
                if (IncludeProperties)
                {
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
                }

                string[] HeaderArray = Header.Split(';');
                List<NomenkProperties> nProperties = new List<NomenkProperties>();

                //connection.Open();
                string Line;
                string Value;
                string NomenkID;

                reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        NomenkProperties np = new NomenkProperties();
                        np.NomenkID = reader.GetString(0).Trim();
                        np.Property = reader.GetString(1).Trim();
                        if (!reader.IsDBNull(2))
                            np.Data = reader.GetString(2).Trim();
                        else
                            np.Data = "";

                        nProperties.Add(np);
                    }
                    if (reader.NextResult())
                    {
                        using (MemoryStream stream = new MemoryStream())
                        {
                            using (TextWriter tw = new StreamWriter(stream, System.Text.Encoding.UTF8))
                            {
                                tw.WriteLine(Header);
                                List<string> UsedNomID = new List<string>();
                                while (reader.Read())
                                {
                                    NomenkID = reader.GetString(0).Trim();
                                    if (!UsedNomID.Contains(NomenkID))
                                    {
                                        UsedNomID.Add(NomenkID);
                                        if (reader.GetString(12).Trim() == "1") //флаг Web
                                            continue;
                                        double Ostatok = Convert.ToDouble(reader.GetValue(11), CultureInfo.InvariantCulture) -
                                            Convert.ToDouble(reader["rezerv"], CultureInfo.InvariantCulture) -
                                            Convert.ToDouble(reader["avSp"], CultureInfo.InvariantCulture);  //Остаток Экран
                                        double OstatokDS = Convert.ToDouble(reader["OstatokDS"], CultureInfo.InvariantCulture) -
                                            Convert.ToDouble(reader["RezervDS"], CultureInfo.InvariantCulture);  //Остаток ДомСад
                                        if ((Convert.ToDouble(reader.GetValue(20), CultureInfo.InvariantCulture) == 1) && (Ostatok + OstatokDS <= 0)) //флаг Снят с производства
                                            continue;
                                        Line = reader.GetString(1).Trim() + ";";
                                        Line = Line + CheckForQuotes(reader.GetString(2).Trim()) + ";";
                                        if (!reader.IsDBNull(3)) //номенклатура характеристики
                                        {
                                            Value = reader.GetString(3).Trim();
                                            if (Value.Length > 0)
                                            {
                                                Value = System.Text.RegularExpressions.Regex.Replace(Value, @"\t|\n|\r", "");
                                                Line = Line + CheckForQuotes(Value);
                                            }
                                        }
                                        Line = Line + ";";
                                        if (!reader.IsDBNull(4)) //производитель
                                        {
                                            Value = reader.GetString(4).Trim();
                                            if ((!reader.IsDBNull(5)) && (reader.GetString(5).Trim() != ""))
                                                Value = reader.GetString(5).Trim();
                                            Line = Line + CheckForQuotes(Value);
                                        }
                                        Line = Line + ";";
                                        if (!reader.IsDBNull(6))  //артикул
                                            Line = Line + CheckForQuotes(reader.GetString(6).Trim());
                                        if (IncludeProperties)
                                        {
                                            for (int i = 6; i <= 14; i++)
                                            {
                                                Value = (nProperties.FirstOrDefault(n => n.NomenkID == NomenkID && n.Property == HeaderArray[i - 1]) ?? new NomenkProperties { Data = "" }).Data;
                                                Line = Line + ";" + CheckForQuotes(Value);
                                            }
                                        }
                                        Line = Line + ";";
                                        Value = reader.GetString(7).Trim(); //Страна
                                        if (String.IsNullOrEmpty(Value))
                                            Value = "Россия";
                                        else if (Value[0] == '-')
                                            Value = "Россия";
                                        else
                                            Value = CheckForQuotes(Value);
                                        Line = Line + Value + ";";

                                        string Gr1 = RemoveLeadingNumbers(reader.GetString(8).Trim());
                                        string Gr2 = RemoveLeadingNumbers(reader.GetString(9).Trim());
                                        string Gr3 = RemoveLeadingNumbers(reader.GetString(10).Trim());

                                        if (!String.IsNullOrEmpty(Gr1))
                                        {
                                            Line = Line + CheckForQuotes(Gr1) + ";"; //Группа уровень 1
                                            Line = Line + CheckForQuotes(Gr2) + ";"; //Группа уровень 2
                                            Line = Line + CheckForQuotes(Gr3) + ";"; //Группа уровень 3
                                        }
                                        else if (!String.IsNullOrEmpty(Gr2))
                                        {
                                            Line = Line + CheckForQuotes(Gr2) + ";"; //Группа уровень 1
                                            Line = Line + CheckForQuotes(Gr3) + ";"; //Группа уровень 2
                                            Line = Line + ";"; //Группа уровень 3
                                        }
                                        else if (!String.IsNullOrEmpty(Gr3))
                                        {
                                            Line = Line + CheckForQuotes(Gr3) + ";"; //Группа уровень 1
                                            Line = Line + ";"; //Группа уровень 2
                                            Line = Line + ";"; //Группа уровень 3
                                        }
                                        else
                                        {
                                            Line = Line + ";"; //Группа уровень 1
                                            Line = Line + ";"; //Группа уровень 2
                                            Line = Line + ";"; //Группа уровень 3
                                        }

                                        if (Ostatok + OstatokDS > 0)
                                            Line = Line + (Ostatok > 0 ? "1;" : "4;");
                                        else
                                            if (Convert.ToDouble(reader["SP10397"], CultureInfo.InvariantCulture) == 0) //флаг НеЗакупать
                                        {
                                            if (Convert.ToDouble(reader["ostPrih"], CultureInfo.InvariantCulture) > 0)
                                                Line = Line + "0;";
                                            else
                                                Line = Line + "3;";
                                        }
                                        else
                                            Line = Line + "2;";
                                        double Opt = Convert.ToDouble(reader.GetValue(13), CultureInfo.InvariantCulture); //Оптовая
                                        double OptSP = Convert.ToDouble(reader.GetValue(14), CultureInfo.InvariantCulture); //Оптовая СП
                                        double Rozn = Convert.ToDouble(reader.GetValue(15), CultureInfo.InvariantCulture); //Розничная
                                        double RoznSP = Convert.ToDouble(reader.GetValue(16), CultureInfo.InvariantCulture); //Розничная СП
                                        //if (reader.GetString(12).Trim() == "1") //флаг Web
                                        //{
                                        //    Opt = 0;
                                        //    OptSP = 0;
                                        //}
                                        Line = Line + ((Rozn > 0) ? Rozn.ToString() : "") + ";" +
                                            ((RoznSP > 0) ? RoznSP.ToString() : "") + ";" +
                                            ((Opt > 0) ? Opt.ToString() : "") + ";" +
                                            ((OptSP > 0) ? OptSP.ToString() : "") + ";" +
                                            ((Rozn > 0) ? "RUB" : "") + ";" +
                                            ((RoznSP > 0) ? "RUB" : "") + ";" +
                                            ((Opt > 0) ? "RUB" : "") + ";" +
                                            ((OptSP > 0) ? "RUB" : "") + ";";
                                        double MinRozn = 0;
                                        double MinOpt = 0;
                                        double SkidkaRozn = 0;
                                        double SkidkaOpt = 0;
                                        if ((Rozn > 0) && (RoznSP > 0))
                                        {
                                            MinRozn = RoznSP;
                                            SkidkaRozn = (Rozn - RoznSP) * 100 / Rozn;
                                        }
                                        else if (Rozn > 0)
                                        {
                                            MinRozn = Rozn;
                                        }
                                        if ((Opt > 0) && (OptSP > 0))
                                        {
                                            MinOpt = Math.Min(Opt, OptSP);
                                            SkidkaOpt = (Opt - OptSP) * 100 / Opt;
                                        }
                                        else if (Opt > 0)
                                        {
                                            MinOpt = Opt;
                                        }
                                        Line = Line + ((MinRozn > 0) ? MinRozn.ToString() : "") + ";" +
                                            ((MinOpt > 0) ? MinOpt.ToString() : "") + ";" +
                                            ((SkidkaRozn > 0) ? SkidkaRozn.ToString() : "") + ";" +
                                            ((SkidkaOpt > 0) ? SkidkaOpt.ToString() : "");
                                        double KolUpak = Convert.ToDouble(reader["kolUpak"], CultureInfo.InvariantCulture); //Количество к перемещению опт
                                        if (KolUpak < 1)
                                            KolUpak = 1;
                                        Line += ";" + KolUpak.ToString();
                                        if (Convert.ToDouble(reader["YandexMarket"], CultureInfo.InvariantCulture) > 0)
                                            Line += ";1";
                                        else
                                            Line += ";0";
                                        if (Convert.ToDouble(reader["TolkoRozn"], CultureInfo.InvariantCulture) > 0)
                                            Line += ";1";
                                        else
                                            Line += ";0";
                                        if (IncludeProperties)
                                        {
                                            for (int i = 32; i <= 110; i++)
                                            {
                                                Value = (nProperties.FirstOrDefault(n => n.NomenkID == NomenkID && n.Property == HeaderArray[i - 1]) ?? new NomenkProperties { Data = "" }).Data;
                                                Line = Line + ";" + CheckForQuotes(Value);
                                            }
                                        }

                                        tw.WriteLine(Line);
                                    }
                                }
                            }
                            reply.ExcelFile = stream.ToArray();
                        }

                        //File.WriteAllBytes(@"f:\tmp\test.csv", reply.ExcelFile);
                    }

                    reply.ReplyState.State = "Success";
                    reply.ReplyState.Description = "Информация успешно извлечена из БД";
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }

            return reply;
        }

        private string SetConditions(List<GroupCondition> conditions, List<string> ListTabName, bool IsFirst = true)
        {
            string q = "";
            string name = "";
            int nGr = 0;
            foreach (GroupCondition gc in conditions)
            {
                if (ListTabName.Contains(gc.name))
                {
                    if (nGr > 0)
                        q = q + " " + gc.conditionType + " ";
                    if ((gc.name == "") && (ListTabName.Contains("rg")))
                        name = "rg";
                    else if ((gc.name == "") && (ListTabName.Contains("ra")))
                        name = "ra";
                    else
                        name = gc.name;

                    if (gc.subCondition != null)
                        q = q + SetConditions(gc.subCondition, ListTabName, false);
                    else
                    {
                        int n = 0;
                        foreach (DDS d in gc.conditionsDDS)
                        {
                            if (n > 0)
                                q = q + " " + d.conditionType + " ";
                            q = q + "(" + ((d.name == "") ? name : d.name) + "." + d.param + d.condition + ")";
                            n++;
                        }
                        if (n > 1)
                            q = "(" + q + ")";
                    }
                    nGr++;
                }
            }
            if (nGr > 1)
                q = "(" + q + ")";
            if ((!String.IsNullOrEmpty(q)) && (IsFirst))
                q = " and " + q;
            return q;
        }

        private string ZaprosTextOst(string Reg, int CurNum, int TabNum, string Param1, string Param2, string Param3, List<DDS> Params, List<GroupCondition> Conditions = null, string Cond1 = "", string Cond2 = "", int Type = 0) //Type=0 -- all; =1 -- only rg; =2 -- only ra
        {
            string q = "select ";
            foreach (DDS d in Params)
            {
                if (String.IsNullOrEmpty(d.param))
                    q = q + "0 as " + d.name + ",";
                else
                    q = q + d.name + ",";
            }
            for (int i = 0; i <= TabNum; i++)
            {
                q = q + "prih" + i.ToString() + ",rash" + i.ToString();
                if (i < TabNum)
                    q = q + ",";
            }
            q = q + " from (select ";
            if (Type != 2)
            {
                foreach (DDS d in Params)
                {
                    if (String.IsNullOrEmpty(d.param))
                        q = q + "0 as " + d.name + ",";
                    else
                        q = q + "rg." + d.param + " as " + d.name + ",";
                }
                for (int i = 0; i <= TabNum; i++)
                {
                    q = q + "0 as prih" + i.ToString() + ",0 as rash" + i.ToString();
                    if (i < TabNum)
                        q = q + ",";
                }
                q = q + " from rg" + Reg + " as rg (NOLOCK) ";
                if (!String.IsNullOrEmpty(Cond1))
                    q = q + Cond1 + " ";
                q = q + "where rg.period = @" + Param1;
                if (Conditions != null)
                    q = q + SetConditions(Conditions, new List<string> { "", "rg" });
            }
            if (Type == 0)
            {
                q = q + " union all select ";
            }
            if (Type != 1)
            {
                bool NeedComma = false;
                string CurParam = "";
                foreach (DDS d in Params)
                {
                    if (NeedComma)
                        q = q + ",";
                    else
                        NeedComma = true;
                    if (d.name == "kolvo" + CurNum.ToString())
                    {
                        CurParam = d.param;
                        q = q + "0 as " + d.name;
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(d.param))
                            q = q + "0 as " + d.name;
                        else
                            q = q + "ra." + d.param + " as " + d.name;
                    }
                }
                for (int i = 0; i <= TabNum; i++)
                {
                    if (i != CurNum)
                        q = q + ",0 as prih" + i.ToString() + ", 0 as rash" + i.ToString();
                    else
                        q = q + ", (ra." + CurParam + "*((ra.debkred+1)%2)) as prih" + CurNum.ToString() + ", (ra." + CurParam + "*ra.debkred) as rash" + CurNum.ToString();
                }
                q = q + " from ra" + Reg + " as ra (NOLOCK) inner join _1sjourn as j (NOLOCK) on (ra.iddoc = j.iddoc) ";
                if (!String.IsNullOrEmpty(Cond2))
                    q = q + Cond2 + " ";
                q = q + @"where j.date_time_iddoc >= convert(varchar(8), @" + Param2 + @", 112) and j.date_time_iddoc <= convert(varchar(8), @" + Param3 + @", 112)";
                if (Conditions != null)
                    q = q + SetConditions(Conditions, new List<string> { "", "ra", "j" });
            }
            q = q + ") x ";

            return q;
        }

        private string ZaprosTextOstMiddleMonth(string Reg, int CurNum, int TabNum, string Param1, string Param2, string Param3, string Param4, string Param5, List<DDS> Params, List<GroupCondition> Conditions = null, string Cond1 = "", string Cond2 = "", bool OnlyMove = false)
        {
            string q = "select ";
            foreach (DDS d in Params)
            {
                if (String.IsNullOrEmpty(d.param))
                    q = q + "0 as " + d.name + ",";
                else
                    q = q + d.name + ",";
            }
            for (int i = 0; i <= TabNum; i++)
            {
                q = q + "prih" + i.ToString() + ",rash" + i.ToString();
                if (i < TabNum)
                    q = q + ",";
            }
            q = q + " from (select ";
            bool NeedComma = false;
            if (!OnlyMove)
            {
                foreach (DDS d in Params)
                {
                    if (d.name == "kolvo" + CurNum.ToString())
                        q = q + "(kolvoS + P1 - R1) as " + d.name + ",";
                    else
                        if (String.IsNullOrEmpty(d.param))
                        q = q + "0 as " + d.name + ",";
                    else
                        q = q + d.name + ",";
                }
                for (int i = 0; i <= TabNum; i++)
                {
                    q = q + "0 as prih" + i.ToString() + ",0 as rash" + i.ToString();
                    if (i < TabNum)
                        q = q + ",";
                }
                q = q + " from (select ";
                NeedComma = false;
                foreach (DDS d in Params)
                {
                    if (NeedComma)
                        q = q + ",";
                    else
                        NeedComma = true;
                    if (d.name == "kolvo" + CurNum.ToString())
                        q = q + d.param + " as kolvoS, 0 as P1, 0 as R1";
                    else
                        if (String.IsNullOrEmpty(d.param))
                        q = q + "0 as " + d.name;
                    else
                        q = q + "rg." + d.param + " as " + d.name;
                }
                q = q + " from rg" + Reg + " as rg (NOLOCK) ";
                if (!String.IsNullOrEmpty(Cond1))
                    q = q + Cond1 + " ";
                q = q + "where rg.period = @" + Param1;
                if (Conditions != null)
                    q = q + SetConditions(Conditions, new List<string> { "", "rg" });
                q = q + " union all select ";
                NeedComma = false;
                foreach (DDS d in Params)
                {
                    if (NeedComma)
                        q = q + ",";
                    else
                        NeedComma = true;
                    if (d.name == "kolvo" + CurNum.ToString())
                        q = q + "0 as kolvoS,(ra." + d.param + "*((ra.debkred+1)%2)) as P1, (ra." + d.param + "*ra.debkred) as R1";
                    else
                        if (String.IsNullOrEmpty(d.param))
                        q = q + "0 as " + d.name;
                    else
                        q = q + "ra." + d.param + " as " + d.name;
                }
                q = q + " from ra" + Reg + @" as ra (NOLOCK)
                    inner join _1sjourn as j (NOLOCK) on (ra.iddoc = j.iddoc) ";
                if (!String.IsNullOrEmpty(Cond2))
                    q = q + Cond2 + " ";
                q = q + "where j.date_time_iddoc >= convert(varchar(8), @" + Param2 + @", 112) and j.date_time_iddoc <= convert(varchar(8), @" + Param3 + @", 112)";
                if (Conditions != null)
                    q = q + SetConditions(Conditions, new List<string> { "", "ra", "j" });
                q = q + ") x ";
                q = q + "union all select ";
            }
            NeedComma = false;
            string CurParam = "";
            foreach (DDS d in Params)
            {
                if (NeedComma)
                    q = q + ",";
                else
                    NeedComma = true;
                if (d.name == "kolvo" + CurNum.ToString())
                {
                    CurParam = d.param;
                    q = q + "0 as " + d.name;
                }
                else
                    if (String.IsNullOrEmpty(d.param))
                    q = q + "0 as " + d.name;
                else
                    q = q + "ra." + d.param + " as " + d.name;
            }
            for (int i = 0; i <= TabNum; i++)
            {
                if (i != CurNum)
                    q = q + ",0 as prih" + i.ToString() + ", 0 as rash" + i.ToString();
                else
                    q = q + ", (ra." + CurParam + "*((ra.debkred+1)%2)) as prih" + CurNum.ToString() + ", (ra." + CurParam + "*ra.debkred) as rash" + CurNum.ToString();
            }
            q = q + " from ra" + Reg + " as ra (NOLOCK) inner join _1sjourn as j (NOLOCK) on (ra.iddoc = j.iddoc) ";
            if (!String.IsNullOrEmpty(Cond2))
                q = q + Cond2 + " ";
            q = q + "where j.date_time_iddoc >= convert(varchar(8), @" + Param4 + @", 112) and j.date_time_iddoc <= convert(varchar(8), @" + Param5 + @", 112)";
            if (Conditions != null)
                q = q + SetConditions(Conditions, new List<string> { "", "ra", "j" });
            q = q + ") x ";

            return q;
        }

        private string ZaprosTZ(DateTime startDate, bool IsTA)
        {
            string q = @"from (
            select nomk, Sum(regTZ) as regTZ, Sum(ostatok) as ostatok
            from (
                select SP10832 as nomk, SP10834 as regTZ, 0 as Ostatok
                from SC10836 (NOLOCK)
                where parentext = '" + PodSklad.TorgZal + @"' and IsMark = 0
                union all ";

            if (IsTA)
            {
                q = q + @"SELECT SP408 As nomk, 0 as regTZ, SP411 As Ostatok
                    FROM
                        RG405 (NOLOCK)
                    WHERE
                        PERIOD = @Period and SP9139 = '" + PodSklad.TorgZal + @"' and
                        SP4062 in " + OwnFirms.InCondition;
            }
            else
            {
                q = q + @"SELECT TMP.nomk AS nomk, 0 as regTZ,
                        SUM(TMP.ostN + TMP.pr - TMP.rh) As Ostatok
                    FROM (
                        SELECT
                            rg.SP408 As nomk,
                            rg.SP411 As ostN,
                            0 As pr,
                            0 As rh
                        FROM
                            RG405 As rg (NOLOCK)
                        WHERE
                            rg.PERIOD = @PeriodR and SP9139 = '" + PodSklad.TorgZal + @"' and
                            rg.SP4062 in " + OwnFirms.InCondition;
                if (startDate != new DateTime(startDate.Year, startDate.Month, 1))
                    q = q + @"
                        UNION ALL
                        SELECT
                            ra1.SP408 As nomk,
                            (ra1.SP411*((ra1.DEBKRED+1)%2))- (ra1.SP411*ra1.DEBKRED) As ostN,
                            0 As pr,
                            0 As rh
                        FROM
                            RA405 As ra1 (NOLOCK)
                        INNER JOIN _1SJOURN As TabJ (NOLOCK)
                            ON (ra1.IDDOC = TabJ.IDDOC)
                        WHERE
                            TabJ.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN1, 112) AND TabJ.DATE_TIME_IDDOC < convert(varchar(8), @PeriodN, 112) and ra1.SP9139 = '" + PodSklad.TorgZal + @"' and
                            ra1.SP4062 in " + OwnFirms.InCondition;
                q = q + @"
                        UNION ALL
                        SELECT
                            ra2.SP408 As nomk,
                            0 As ostN,
                            (ra2.SP411*((ra2.DEBKRED+1)%2)) As pr,
                            (ra2.SP411*ra2.DEBKRED) As rh
                        FROM
                            RA405 As ra2 (NOLOCK)
                        INNER JOIN _1SJOURN As TabJ (NOLOCK)
                            ON (ra2.IDDOC = TabJ.IDDOC)
                        WHERE
                            TabJ.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) AND TabJ.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112) and ra2.SP9139 = '" + PodSklad.TorgZal + @"' and
                            ra2.SP4062 in " + OwnFirms.InCondition + @"
                        ) AS tmp GROUP BY nomk";
            }
            q = q + @") as t
            group by nomk) as R ";
            return q;
        }

        private string FixZaprosProdag = @"From (
            SELECT
                tmp.nomk as nomk, tmp.sklad as sklad, min(tmp.dateDoc) as MinDocDate, max(tmp.dateDoc) as MaxDocDate,
                sum(tmp.prodano - tmp.vozvrat) as prodano,
                sum(tmp.prodST - tmp.prodSTV) as prStoimost,
                sum(tmp.sebeST - tmp.sebeSTV) as sebest,
                sum(tmp.PoGarantii) as PoGarantii, sum(tmp.PoGarantiiSt) as PoGarantiiSt, sum(tmp.PoGarantiiSebeSt) as PoGarantiiSebeSt,
                sum(tmp.Platno) as Platno, sum(tmp.PlatnoSt) as PlatnoSt, sum(tmp.PlatnoSebeSt) as PlatnoSebeSt
            FROM (
                select
                    ra.SP9585 as nomk,
                    ra.SP9589 as sklad,
                    left(j.DATE_TIME_IDDOC,8) as dateDoc,
                    (ra.SP9592*((ra.debkred+1)%2)) as prodano, (ra.SP9595*((ra.debkred+1)%2)) as vozvrat,
                    (ra.SP9591*((ra.debkred+1)%2)) as prodST, (ra.SP9594*((ra.debkred+1)%2)) as prodSTV,
                    (ra.SP9590*((ra.debkred+1)%2)) as sebeST, (ra.SP9593*((ra.debkred+1)%2)) as sebeSTV,
                    0 as PoGarantii, 0 as PoGarantiiSt, 0 as PoGarantiiSebeSt,
                    0 as Platno, 0 as PlatnoSt, 0 as PlatnoSebeSt
                from
                    ra9596 as ra (NOLOCK)
                inner join _1sjourn as j (NOLOCK) on (j.iddoc = ra.iddoc)
                where
                    j.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) and j.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112)
                    and ra.SP10787<>'   15P   '
                    and ra.SP9588 in " + OwnFirms.InCondition + @" 

                UNION ALL

                select
                    ra.SP10299 as nomk,
                    ra.SP10295 as sklad,
                    left(j.DATE_TIME_IDDOC,8) as dateDoc,
                    0 as prodano, 0 as vozvrat,
                    0 as prodST, 0 as prodSTV,
                    0 as sebeST, 0 as sebeSTV,
                    (ra.SP10304*((ra.debkred+1)%2)) as PoGarantii, (ra.SP10303*((ra.debkred+1)%2)) as PoGarantiiSt, (ra.SP10302*((ra.debkred+1)%2)) as PoGarantiiSebeSt,
                    0 as Platno, 0 as PlatnoSt, 0 as PlatnoSebeSt
                from
                    ra10305 as ra (NOLOCK)
                inner join _1sjourn as j (NOLOCK) on (j.iddoc = ra.iddoc)
                where
                    j.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) and j.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112)
                    and ra.SP10301<>0
                    and ra.SP10294 in " + OwnFirms.InCondition + @" 

                UNION ALL

                select
                    ra.SP10299 as nomk,
                    ra.SP10295 as sklad,
                    left(j.DATE_TIME_IDDOC,8) as dateDoc,
                    0 as prodano, 0 as vozvrat,
                    0 as prodST, 0 as prodSTV,
                    0 as sebeST, 0 as sebeSTV,
                    0 as PoGarantii, 0 as PoGarantiiSt, 0 as PoGarantiiSebeSt,
                    (ra.SP10304*((ra.debkred+1)%2)) as Platno, (ra.SP10303*((ra.debkred+1)%2)) as PlatnoSt, (ra.SP10302*((ra.debkred+1)%2)) as PlatnoSebeSt
                from
                    ra10305 as ra (NOLOCK)
                inner join _1sjourn as j (NOLOCK) on (j.iddoc = ra.iddoc)
                where
                    j.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) and j.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112)
                    and ra.SP10301=0
                    and ra.SP10294 in " + OwnFirms.InCondition + @" 

                UNION ALL

                select
                    ra.SP408 as nomk,
                    ra.SP418 as sklad,
                    left(j.DATE_TIME_IDDOC,8) as dateDoc,
                    0 as prodano, 0 as vozvrat,
                    0 as prodST, 0 as prodSTV,
                    0 as sebeST, 0 as sebeSTV,
                    0 as PoGarantii, 0 as PoGarantiiSt, 0 as PoGarantiiSebeSt,
                    0 as Platno, 0 as PlatnoSt, 0 as PlatnoSebeSt
                from
                    ra405 as ra (NOLOCK)
                inner join _1sjourn as j (NOLOCK) on (j.iddoc = ra.iddoc)
                where
                    j.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) and j.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112)

            ) as tmp
            group by nomk,sklad) as R ";

        private string ZaprosOstatkiSklad(DateTime startDate, DateTime endDate, bool IsTA = false)
        {
            string q = "";
            int tableN = 2;

            if (!IsTA)
            {
                if (startDate == new DateTime(startDate.Year, startDate.Month, 1))
                {
                    q = "From (select nomk, sklad, Sum(kolvo0) as NachOst, Sum(prih0) as Prihod, Sum(rash0) as Rashod, Sum(kolvo0 + prih0 - rash0) as KonOst, " +
                         "Sum(kolvo1) as NabN, Sum(kolvo1 + prih1 - rash1) as NabK, Sum(kolvo2) as RezN, Sum(kolvo2 + prih2 - rash2) as RezK from (";
                    q = q + ZaprosTextOst("405", 0, tableN, "PeriodR", "PeriodN", "PeriodK", new List<DDS>() { new DDS() { param = "sp408", name = "nomk" },
                                                                                                                new DDS() { param = "sp418", name = "sklad"},
                                                                                                                new DDS() { param = "sp411", name = "kolvo0"},
                                                                                                                new DDS() { param = "", name = "kolvo1"},
                                                                                                                new DDS() { param = "", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4062", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                );
                    q = q + "union all " + ZaprosTextOst("11973", 1, tableN, "PeriodR", "PeriodN", "PeriodK", new List<DDS>() {    new DDS() { param = "SP11971", name = "nomk" },
                                                                                                                            new DDS() { param = "SP11967", name = "sklad"},
                                                                                                                            new DDS() { param = "", name="kolvo0"},
                                                                                                                            new DDS() { param = "SP11972", name = "kolvo1"},
                                                                                                                            new DDS() { param = "", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11966", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                            );
                    q = q + "union all " + ZaprosTextOst("4480", 2, tableN, "PeriodR", "PeriodN", "PeriodK", new List<DDS>() {    new DDS() { param = "SP4477", name = "nomk" },
                                                                                                                            new DDS() { param = "SP4476", name = "sklad"},
                                                                                                                            new DDS() { param = "", name="kolvo0"},
                                                                                                                            new DDS() { param = "", name = "kolvo1"},
                                                                                                                            new DDS() { param = "SP4479", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4475", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                            );
                    q = q + ") x group by nomk,sklad) as R ";
                }
                else
                {
                    q = "From (select nomk, sklad, Sum(kolvo0) as NachOst, Sum(prih0) as Prihod, Sum(rash0) as Rashod, Sum(kolvo0 + prih0 - rash0) as KonOst, " +
                        "Sum(kolvo1) as NabN, Sum(kolvo1 + prih1 - rash1) as NabK, Sum(kolvo2) as RezN, Sum(kolvo2 + prih2 - rash2) as RezK from (" +
                        ZaprosTextOstMiddleMonth("405", 0, tableN, "PeriodR", "PeriodN1", "PeriodK1", "PeriodN", "PeriodK", new List<DDS>() {
                            new DDS() { param = "sp408", name = "nomk"},
                            new DDS() { param = "sp418", name = "sklad"},
                            new DDS() { param = "sp411", name = "kolvo0"},
                            new DDS() { param = "", name = "kolvo1"},
                            new DDS() { param = "", name = "kolvo2"}
                        },
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4062", condition = " in " + OwnFirms.InCondition } } } }
                        ) + "union all " +
                        ZaprosTextOstMiddleMonth("11973", 1, tableN, "PeriodR", "PeriodN1", "PeriodK1", "PeriodN", "PeriodK", new List<DDS>() {
                            new DDS() { param = "SP11971", name = "nomk"},
                            new DDS() { param = "SP11967", name = "sklad"},
                            new DDS() { param = "", name = "kolvo0"},
                            new DDS() { param = "SP11972", name = "kolvo1"},
                            new DDS() { param = "", name = "kolvo2"}
                        },
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11966", condition = " in " + OwnFirms.InCondition } } } }
                        ) + "union all " +
                        ZaprosTextOstMiddleMonth("4480", 2, tableN, "PeriodR", "PeriodN1", "PeriodK1", "PeriodN", "PeriodK", new List<DDS>() {
                            new DDS() { param = "SP4477", name = "nomk"},
                            new DDS() { param = "SP4476", name = "sklad"},
                            new DDS() { param = "", name = "kolvo0"},
                            new DDS() { param = "", name = "kolvo1"},
                            new DDS() { param = "SP4479", name = "kolvo2"}
                        },
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4475", condition = " in " + OwnFirms.InCondition } } } }
                        );
                    q = q + ") x group by nomk,sklad) as R ";

                }
            }
            else
            {

                q = @"From (select nomk, sklad, Sum(kolvo0 - prih0 + rash0) as NachOst, Sum(prih0) as Prihod, Sum(rash0) as Rashod, Sum(kolvo0) as KonOst, Sum(kolvo1 - prih1 + rash1) as NabN, Sum(kolvo1) as NabK, Sum(kolvo2 - prih2 + rash2) as RezN, Sum(kolvo2) as RezK from 
                                (" + ZaprosTextOst("405", 0, tableN, "Period", "PeriodN", "PeriodK", new List<DDS>() { new DDS() { param = "sp408", name = "nomk" },
                                                                                                                new DDS() { param = "sp418", name = "sklad"},
                                                                                                                new DDS() { param = "sp411", name = "kolvo0"},
                                                                                                                new DDS() { param = "", name = "kolvo1"},
                                                                                                                new DDS() { param = "", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4062", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                );
                q = q + "union all " + ZaprosTextOst("11973", 1, tableN, "Period", "PeriodN", "PeriodK", new List<DDS>() {    new DDS() { param = "SP11971", name = "nomk" },
                                                                                                                        new DDS() { param = "SP11967", name = "sklad"},
                                                                                                                        new DDS() { param = "", name = "kolvo0"},
                                                                                                                        new DDS() { param = "SP11972", name = "kolvo1"},
                                                                                                                        new DDS() { param = "", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11966", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                        );
                q = q + "union all " + ZaprosTextOst("4480", 2, tableN, "Period", "PeriodN", "PeriodK", new List<DDS>() {    new DDS() { param = "SP4477", name = "nomk" },
                                                                                                                            new DDS() { param = "SP4476", name = "sklad"},
                                                                                                                            new DDS() { param = "", name="kolvo0"},
                                                                                                                            new DDS() { param = "", name = "kolvo1"},
                                                                                                                            new DDS() { param = "SP4479", name = "kolvo2"}},
                            new List<GroupCondition>() { new GroupCondition() { name = "", conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP4475", condition = " in " + OwnFirms.InCondition } } } }
                                                                                                                            );
                q = q + ") x group by nomk,sklad) as R ";
            }

            return q;
        }

        private string PrepareZaprosDop(DateTime startDate, bool IsTA = false)
        {
            int totalTableNumber = 9;
            List<TableData> ListTableData = new List<TableData>();

            int curTable = 0;
            TableData td = new TableData();
            td.reg = "8696"; //ОстаткиДоставки
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP8698", name = "nomk" });
            td.paramsDDS.Add(new DDS() { param = "SP8701", name = "kolvo0" });
            for (int i = 1; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>();
            GroupCondition gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP11041", name = "", condition = "=0" } }; //Условие (ЭтоИзделие = 0)
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP8697", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "351"; //Комиссия
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP354", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP357", name = "kolvo1" });
            for (int i = 2; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP4063", name = "", condition = " in " + OwnFirms.InCondition }
            }
            } };
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "9981"; //ОстаткиНаИзделиях
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP9978", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP9979", name = "kolvo2" });
            for (int i = 3; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP9973", name = "", condition = " in " + OwnFirms.InCondition }
            }
            } };
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "9972"; //ПартииМастерской
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP9960", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP9970", name = "kolvo3" });
            for (int i = 4; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP9964", name = "", condition = "='     0   '" } }; //Условие(ПустоеЗначение(Контрагент)=1)
            td.conditionsGroup.Add(gc);
            ListTableData.Add(td);

            //ПриходВнешний начало
            curTable++;
            td = new TableData();
            td.reg = "405"; //ОстаткиТМЦ
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo4" });
            for (int i = 5; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.Cond2 = "left join DH1582 (NOLOCK) on (DH1582.IDDOC = j.IDDOC)";
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.subCondition = new List<GroupCondition>();

            GroupCondition gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=1582" }, //(ТекДок.Вид()=""ПоступлениеТМЦ"")
                new DDS() { name = "DH1582", param = "SP8996", condition = "='          '", conditionType = "and" } //ПустоеЗначение(ТекДок.НомерРеализации) = 1
            };
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=3957" } }; //(ТекДок.Вид()=""ВводОстатковТМЦ"")
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=9040" } }; //(ТекДок.Вид()=""АктПриемкиТовараПриход"")
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=2106" } }; //(ТекДок.Вид()=""ОприходованиеТМЦ"")
            gc.subCondition.Add(gc1);

            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);

            td = new TableData();
            td.reg = "11973"; //НаборНаСкладе
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo4" });
            for (int i = 5; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.Cond2 = "left join DH1582 (NOLOCK) on (DH1582.IDDOC = j.IDDOC)";
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.subCondition = new List<GroupCondition>();

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=1582" }, //(ТекДок.Вид()=""ПоступлениеТМЦ"")
                new DDS() { name = "DH1582", param = "SP8996", condition = "='          '", conditionType = "and" } //ПустоеЗначение(ТекДок.НомерРеализации) = 1
            };
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=3957" } }; //(ТекДок.Вид()=""ВводОстатковТМЦ"")
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=9040" } }; //(ТекДок.Вид()=""АктПриемкиТовараПриход"")
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=2106" } }; //(ТекДок.Вид()=""ОприходованиеТМЦ"")
            gc.subCondition.Add(gc1);

            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);
            //ПриходВнешний конец

            //РасходВнутренний начало
            curTable++;
            td = new TableData();
            td.reg = "405"; //ОстаткиТМЦ
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo5" });
            for (int i = curTable + 1; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=1790" } }; //(ТекДок.Вид()=""СписаниеТМЦ"")
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);

            td = new TableData();
            td.reg = "11973"; //НаборНаСкладе
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo5" });
            for (int i = curTable + 1; i < totalTableNumber; i++)
            {
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            }
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=1790" } }; //(ТекДок.Вид()=""СписаниеТМЦ"")
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);
            //РасходВнутренний конец

            curTable++;
            td = new TableData();
            td.reg = "328"; //ПартииНаличие
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP331", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP342", name = "kolvo6" });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.Cond2 = @"left join DH1611 (NOLOCK) on (DH1611.IDDOC = j.IDDOC)
                         left join DH1774 (NOLOCK) on (DH1774.IDDOC = j.IDDOC)";

            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.subCondition = new List<GroupCondition>();

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=3114" }, //(ТекДок.Вид()=""ОтчетККМ"")
                new DDS() { name = "ra", param = "SP347", condition = "='   3AA   '", conditionType = "and" } //(КодОперацииП = Перечисление.КодыОпераций.РозничнаяПродажаЕНВД)
            };
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=3114" }, //(ТекДок.Вид()=""ОтчетККМ"")
                new DDS() { name = "ra", param = "SP347", condition = "='   2GU   '", conditionType = "and" } //(КодОперацииП = Перечисление.КодыОпераций.РозничнаяПродажа)
            };
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=1611" }, //(ТекДок.Вид()=""Реализация"")
                new DDS() { name = "DH1611", param = "SP1583", condition = " not in ('   19LS  ','    15S  ','    E0S  ')", conditionType = "and" } //(НЕ(ТекДок.Контрагент в СписокСвоихКонтрагентов))
            };
            gc.subCondition.Add(gc1);

            gc1 = new GroupCondition();
            gc1.name = "j";
            gc1.conditionType = "or";
            gc1.conditionsDDS = new List<DDS>() {
                new DDS() { name = "j", param = "IDDOCDEF", condition = "=1774" }, //(ТекДок.Вид()=""ОтчетКомиссионера"")
                new DDS() { name = "DH1774", param = "SP1751", condition = " not in ('   19LS  ','    15S  ','    E0S  ')", conditionType = "and" } //(НЕ(ТекДок.Контрагент в СписокСвоихКонтрагентов))
            };
            gc.subCondition.Add(gc1);

            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP4061", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "405"; //ОстаткиТМЦ
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo7" });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });

            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "j";
            gc.conditionsDDS = new List<DDS>() { new DDS() { name = "j", param = "IDDOCDEF", condition = "=12345" } }; //(ТекДок.Вид()=""ЗаменаНоменклатуры"")
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "12406"; //ГарантийноеЗавершение
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP12403", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP12405", name = "kolvo8" });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP12399", name = "", condition = " in " + OwnFirms.InCondition }
            }
            } };
            ListTableData.Add(td);

            return ZaprosDop(ListTableData, startDate, IsTA);
        }

        private string ZaprosDop(List<TableData> ListTableData, DateTime startDate, bool IsTA = false)
        {
            string q = "";

            if (!IsTA)
            {
                q = @"From
                                (select nomk, Sum(kolvo0) as NachOst, Sum(prih0) as Prihod, Sum(rash0) as Rashod, Sum(kolvo0 + prih0 - rash0) as KonOst,
                                    Sum(kolvo1) as NachKommis, Sum(kolvo1 + prih1 - rash1) as KonKommis,
                                    Sum(kolvo2) as NachUMastera, Sum(kolvo2 + prih2 - rash2) as KonUMastera,
                                    Sum(kolvo3) as NachMastersk, Sum(kolvo3 + prih3 - rash3) as KonMastersk,
                                    Sum(prih4) as PrihVnesh, Sum(rash5) as RashVnutr, Sum(rash6) as RashPartii, Sum(prih7) as ZamenaPlus, Sum(rash7) as ZamenaMinus,
                                    Sum(kolvo8) as NachGarZaver, Sum(kolvo8 + prih8 - rash8) as KonGarZaver
                                from 
                                (";
                if (startDate == new DateTime(startDate.Year, startDate.Month, 1))
                    for (int i = 0; i <= ListTableData.Count - 1; i++)
                    {
                        if (i > 0)
                            q = q + "union all ";
                        q = q + ZaprosTextOst(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "PeriodR", "PeriodN", "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, ListTableData[i].Type);
                    }
                else
                    for (int i = 0; i <= ListTableData.Count - 1; i++)
                    {
                        if (i > 0)
                            q = q + "union all ";
                        q = q + ZaprosTextOstMiddleMonth(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "PeriodR", "PeriodN1", "PeriodK1", "PeriodN", "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, (ListTableData[i].Type == 2));
                    }
            }
            else
            {
                q = @"From (select nomk, Sum(kolvo0 - prih0 + rash0) as NachOst, Sum(prih0) as Prihod, Sum(rash0) as Rashod, Sum(kolvo0) as KonOst, 
                                Sum(kolvo1 - prih1 + rash1) as NachKommis, Sum(kolvo1) as KonKommis,
                                Sum(kolvo2 - prih2 + rash2) as NachUMastera, Sum(kolvo2) as KonUMastera,
                                Sum(kolvo3 - prih3 + rash3) as NachMastersk, Sum(kolvo3) as KonMastersk,
                                Sum(prih4) as PrihVnesh, Sum(rash5) as RashVnutr, Sum(rash6) as RashPartii, Sum(prih7) as ZamenaPlus, Sum(rash7) as ZamenaMinus,
                                Sum(kolvo8 - prih8 + rash8) as NachGarZaver, Sum(kolvo8) as KonGarZaver
                            from (";
                for (int i = 0; i <= ListTableData.Count - 1; i++)
                {
                    if (i > 0)
                        q = q + "union all ";
                    q = q + ZaprosTextOst(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "Period", "PeriodN", "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, ListTableData[i].Type);
                }
            }
            q = q + ") x group by nomk) as R ";

            return q;
        }

        private string SetColumnName(bool IsTA, int column, string name)
        {
            if (IsTA)
                return ",Sum(kolvo" + column.ToString() + ") as " + name;
            else
                return ",Sum(kolvo" + column.ToString() + " + prih" + column.ToString() + " - rash" + column.ToString() + ") as " + name;
        }

        private string ZaprosFilial(List<TableData> ListTableData, DateTime startDate, bool IsTA = false, bool IsMaster = false, bool S_PrihRas = true, bool OstBrak = true, bool PP = true, bool BrakMX = true,
                                    bool Roz = true, bool Otstoy = true, bool Akt = true, bool SobNugdy = true, bool Podarky = true, bool Prokat = true, bool Ilin = true)
        {
            string q = "";
            q = "From (select nomk, Sum(prih0) as ProdanoM, Sum(prih1) as VozvratM";
            int TotNum = 2;
            if (IsMaster)
            {
                q = q + ",Sum(prih2) as PoGarantiiM, Sum(prih3) as PlatnoM" +
                        SetColumnName(IsTA, TotNum + 2, "PodZakaz") +
                        SetColumnName(IsTA, TotNum + 3, "PodZakazPlatno");
                TotNum = TotNum + 4;
            }
            q = q + SetColumnName(IsTA, TotNum, "PredZakaz") +
                    SetColumnName(IsTA, TotNum + 1, "Spros") +
                    SetColumnName(false, TotNum + 2, "SprosOborot") +
                    SetColumnName(IsTA, TotNum + 3, "RezRaspred") +
                    SetColumnName(IsTA, TotNum + 4, "VNabore") +
                    SetColumnName(IsTA, TotNum + 5, "VZayavke") +
                    SetColumnName(IsTA, TotNum + 6, "OPrihod");
            TotNum = TotNum + 7;
            if (OstBrak)
            {
                q = q + SetColumnName(IsTA, TotNum, "Brak");
                TotNum++;
            }
            if (!S_PrihRas)
            {
                q = q + SetColumnName(IsTA, TotNum, "Dostavka") +
                        SetColumnName(IsTA, TotNum + 1, "Komis") +
                        SetColumnName(IsTA, TotNum + 2, "UMastera") +
                        SetColumnName(IsTA, TotNum + 3, "VMastersk");
                TotNum = TotNum + 4;
            }
            if (PP)
            {
                q = q + SetColumnName(IsTA, TotNum, "PpBrak");
                TotNum++;
            }
            if (BrakMX)
            {
                q = q + SetColumnName(IsTA, TotNum, "BrakMX");
                TotNum++;
            }
            if (Roz)
            {
                q = q + SetColumnName(IsTA, TotNum, "Roz");
                TotNum++;
            }
            if (Otstoy)
            {
                q = q + SetColumnName(IsTA, TotNum, "Otstoy");
                TotNum++;
            }
            if (Akt)
            {
                q = q + SetColumnName(IsTA, TotNum, "Akt");
                TotNum++;
            }
            if (SobNugdy)
            {
                q = q + SetColumnName(IsTA, TotNum, "SobNugdy");
                TotNum++;
            }
            if (Podarky)
            {
                q = q + SetColumnName(IsTA, TotNum, "Podarky");
                TotNum++;
            }
            if (Prokat)
            {
                q = q + SetColumnName(IsTA, TotNum, "Prokat");
                TotNum++;
            }
            if (Ilin)
            {
                q = q + SetColumnName(IsTA, TotNum, "Ilin");
                TotNum++;
            }
            q = q + " from (";
            if (!IsTA)
            {
                if (startDate == new DateTime(startDate.Year, startDate.Month, 1))
                    for (int i = 0; i <= ListTableData.Count - 1; i++)
                    {
                        if (i > 0)
                            q = q + "union all ";
                        q = q + ZaprosTextOst(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "PeriodR", "PeriodN" + (ListTableData[i].IsMonth ? "M" : (ListTableData[i].IsWeek ? "W" : "")), "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, ListTableData[i].Type);
                    }
                else
                    for (int i = 0; i <= ListTableData.Count - 1; i++)
                    {
                        if (i > 0)
                            q = q + "union all ";
                        q = q + ZaprosTextOstMiddleMonth(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "PeriodR", "PeriodN1", "PeriodK1", "PeriodN" + (ListTableData[i].IsMonth ? "M" : (ListTableData[i].IsWeek ? "W" : "")), "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, (ListTableData[i].Type == 2));
                    }
            }
            else
            {
                for (int i = 0; i <= ListTableData.Count - 1; i++)
                {
                    if (i > 0)
                        q = q + "union all ";
                    q = q + ZaprosTextOst(ListTableData[i].reg, ListTableData[i].curTable, ListTableData[i].totalTable, "Period", "PeriodN" + (ListTableData[i].IsMonth ? "M" : (ListTableData[i].IsWeek ? "W" : "")), "PeriodK", ListTableData[i].paramsDDS, ListTableData[i].conditionsGroup, ListTableData[i].Cond1, ListTableData[i].Cond2, ListTableData[i].Type);
                }
            }
            q = q + ") x group by nomk) as R ";

            return q;
        }

        private string PrepareZaprosFilial(DateTime startDate, bool IsTA = false, bool IsMaster = false, bool S_PrihRas = true, List<string> listSkl = null, List<string> PpBrak = null, List<string> BrakMX = null,
                                            List<string> Roz = null, List<string> Otstoy = null, List<string> Akt = null, List<string> SobNugdy = null, List<string> Podarky = null, List<string> Prokat = null, List<string> Ilin = null)
        {
            int totalTableNumber = 9;
            if ((PpBrak != null) && (PpBrak.Count > 0))
                totalTableNumber++;
            if ((BrakMX != null) && (BrakMX.Count > 0))
                totalTableNumber++;
            if ((Roz != null) && (Roz.Count > 0))
                totalTableNumber++;
            if ((Otstoy != null) && (Otstoy.Count > 0))
                totalTableNumber++;
            if ((Akt != null) && (Akt.Count > 0))
                totalTableNumber++;
            if ((SobNugdy != null) && (SobNugdy.Count > 0))
                totalTableNumber++;
            if ((Podarky != null) && (Podarky.Count > 0))
                totalTableNumber++;
            if ((Prokat != null) && (Prokat.Count > 0))
                totalTableNumber++;
            if ((Ilin != null) && (Ilin.Count > 0))
                totalTableNumber++;
            if ((listSkl != null) && (listSkl.Count > 0))
                totalTableNumber++;
            if (!S_PrihRas)
                totalTableNumber = totalTableNumber + 4;
            if (IsMaster)
                totalTableNumber = totalTableNumber + 4;
            List<TableData> ListTableData = new List<TableData>();

            int curTable = 0;

            TableData td = new TableData();
            td.reg = "9596"; //ПрямыеПродажи за месяц
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP9585", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP9592", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });

            td.conditionsGroup = new List<GroupCondition>();
            GroupCondition gc = new GroupCondition();
            gc.name = "ra";
            gc.conditionsDDS = new List<DDS>() { new DDS() { name = "ra", param = "SP10787", condition = "<>'   15P   '" } }; //(КодОперации <> глКО.ПередачаНаРеализацию)
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP9588", name = "", condition = " in " + OwnFirms.InCondition }
            }
            });
            td.Type = 2;
            td.IsMonth = true;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "9596"; //ПрямыеПродажи за месяц
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP9585", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP9595", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
            {
                new DDS() { param = "SP9588", name = "", condition = " in " + OwnFirms.InCondition }
            }
            } };
            td.IsMonth = true;
            td.Type = 2;
            ListTableData.Add(td);

            if (IsMaster)
            {
                curTable++;
                td = new TableData();
                td.reg = "10305"; //ПродажиМастерской за месяц
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP10299", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP10304", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "ra";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "ra", param = "SP10301", condition = "<>0" } }; //(Гарантия<>0)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP10294", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                td.Type = 2;
                td.IsMonth = true;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "10305"; //ПродажиМастерской за месяц
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP10299", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP10304", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "ra";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "ra", param = "SP10301", condition = "=0" } }; //(Гарантия=0)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP10294", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                td.Type = 2;
                td.IsMonth = true;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "11055"; //СтопЛистЗЧ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11050", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11054", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "11055"; //СтопЛистЗЧ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11050", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11054", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11060", condition = "=0" } }; //(Гарантия=0)
                td.conditionsGroup.Add(gc);
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

            }

            curTable++;
            td = new TableData();
            td.reg = "4667"; //ЗаказыЗаявки
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP4663", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP4666", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP4665", name = "", condition = "='     0   '" } }; //ПустоеЗначение(ЗаказПоставщику)=1
            td.conditionsGroup.Add(gc);
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "12815"; //СпросОстатки
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP12811", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP12814", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP12818", name = "", condition = " in " + OwnFirms.InCondition }
                    }
            } };
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "12791"; //Спрос (обороты)
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP12786", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP12789", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.IsWeek = true;
            td.Type = 2;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "4480"; //РезервыТМЦ
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP4477", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP4479", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP4762", name = "", condition = "='     0   '" } }; //ПустоеЗначение(Заявка)=1
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4475", name = "", condition = " in " + OwnFirms.InCondition }
                    }
            });
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "11973"; //НаборНаСкладе
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            if ((listSkl != null) && (listSkl.Count > 0))
            {
                td.Cond1 = "inner join sc55 (NOLOCK) on (sc55.id = rg.SP11967)";
                td.Cond2 = "inner join sc55 (NOLOCK) on (sc55.id = ra.SP11967)";
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "sc55", param = "code", condition = " in (" + string.Join(",", listSkl.Select(x => "'" + x + "'").ToArray()) + ")" } }; //(Филиал в СписокФилиалов)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
            }
            else
                td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                } };
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "4480"; //РезервыТМЦ
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP4477", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP4479", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP4762", name = "", condition = "<>'     0   '" } }; //ПустоеЗначение(Заявка)=0
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4475", name = "", condition = " in " + OwnFirms.InCondition }
                    }
            });
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            curTable++;
            td = new TableData();
            td.reg = "464"; //Заказы
            td.curTable = curTable;
            td.totalTable = totalTableNumber - 1;
            td.paramsDDS = new List<DDS>();
            td.paramsDDS.Add(new DDS() { param = "SP466", name = "nomk" });
            for (int n = 0; n < td.curTable; n++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
            td.paramsDDS.Add(new DDS() { param = "SP4471", name = "kolvo" + curTable.ToString() });
            for (int i = curTable + 1; i < totalTableNumber; i++)
                td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
            td.conditionsGroup = new List<GroupCondition>();
            gc = new GroupCondition();
            gc.name = "";
            gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP13166", name = "", condition = "<>2" } }; //ТипЗаказа <> 2
            td.conditionsGroup.Add(gc);
            td.conditionsGroup.Add(new GroupCondition()
            {
                name = "",
                conditionType = "and",
                conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4467", name = "", condition = " in " + OwnFirms.InCondition }
                    }
            });
            if (IsTA)
                td.Type = 1;
            ListTableData.Add(td);

            if ((listSkl != null) && (listSkl.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.Cond1 = "inner join sc55 (NOLOCK) on (sc55.id = rg.SP418)";
                td.Cond2 = "inner join sc55 (NOLOCK) on (sc55.id = ra.SP418)";
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "sc55", param = "code", condition = " not in (" + string.Join(",", listSkl.Select(x => "'" + x + "'").ToArray()) + ")" } }; //НЕ(Филиал в СписокФилиалов)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.Cond1 = "inner join sc55 (NOLOCK) on (sc55.id = rg.SP11967)";
                td.Cond2 = "inner join sc55 (NOLOCK) on (sc55.id = ra.SP11967)";
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "sc55", param = "code", condition = " not in (" + string.Join(",", listSkl.Select(x => "'" + x + "'").ToArray()) + ")" } }; //НЕ(Филиал в СписокФилиалов)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if (!S_PrihRas)
            {
                curTable++;
                td = new TableData();
                td.reg = "8696"; //ОстаткиДоставки
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP8698", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP8701", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP11041", name = "", condition = "=0" } }; //Условие (ЭтоИзделие = 0)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP8697", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "351"; //Комиссия
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP354", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP357", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4063", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                } };
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "9981"; //ОстаткиНаИзделиях
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP9978", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP9979", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>() { new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP9973", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                } };
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                curTable++;
                td = new TableData();
                td.reg = "9972"; //ПартииМастерской
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP9960", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP9970", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { param = "SP9964", name = "", condition = "='     0   '" } }; //Условие(ПустоеЗначение(Контрагент)=1)
                td.conditionsGroup.Add(gc);
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((PpBrak != null) && (PpBrak.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", PpBrak.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияПредБрак)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", PpBrak.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияПредБрак)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((BrakMX != null) && (BrakMX.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", BrakMX.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияБрак)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", BrakMX.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияБрак)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Roz != null) && (Roz.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Roz.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияРозыск)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Roz.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияРозыск)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Otstoy != null) && (Otstoy.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Otstoy.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияОтстой)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Otstoy.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияОтстой)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Akt != null) && (Akt.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Akt.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияАкт)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Akt.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияАкт)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((SobNugdy != null) && (SobNugdy.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", SobNugdy.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияСобстНужды)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", SobNugdy.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияСобстНужды)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Podarky != null) && (Podarky.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Podarky.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияПодарки)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Podarky.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияПодарки)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Prokat != null) && (Prokat.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Prokat.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияПрокат)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Prokat.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияПрокат)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            if ((Ilin != null) && (Ilin.Count > 0))
            {
                curTable++;
                td = new TableData();
                td.reg = "405"; //ОстаткиТМЦ
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP408", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP411", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP9139", condition = " in (" + string.Join(",", Ilin.Select(x => "'" + x + "'").ToArray()) + ")" } }; // Когда(ПодСклад в СписокМестХраненияИльин)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP4062", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);

                td = new TableData();
                td.reg = "11973"; //НаборНаСкладе
                td.curTable = curTable;
                td.totalTable = totalTableNumber - 1;
                td.paramsDDS = new List<DDS>();
                td.paramsDDS.Add(new DDS() { param = "SP11971", name = "nomk" });
                for (int n = 0; n < td.curTable; n++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + n.ToString() });
                td.paramsDDS.Add(new DDS() { param = "SP11972", name = "kolvo" + curTable.ToString() });
                for (int i = curTable + 1; i < totalTableNumber; i++)
                    td.paramsDDS.Add(new DDS() { param = "", name = "kolvo" + i.ToString() });
                td.conditionsGroup = new List<GroupCondition>();
                gc = new GroupCondition();
                gc.name = "";
                gc.conditionsDDS = new List<DDS>() { new DDS() { name = "", param = "SP11968", condition = " in (" + string.Join(",", Ilin.Select(x => "'" + x + "'").ToArray()) + ")" } }; //Когда(ПодСклад в СписокМестХраненияИльин)
                td.conditionsGroup.Add(gc);
                td.conditionsGroup.Add(new GroupCondition()
                {
                    name = "",
                    conditionType = "and",
                    conditionsDDS = new List<DDS>()
                    {
                        new DDS() { param = "SP11966", name = "", condition = " in " + OwnFirms.InCondition }
                    }
                });
                if (IsTA)
                    td.Type = 1;
                ListTableData.Add(td);
            }

            return ZaprosFilial(ListTableData, startDate, IsTA, IsMaster, S_PrihRas, ((listSkl != null) && (listSkl.Count > 0)),
                ((PpBrak != null) && (PpBrak.Count > 0)), ((BrakMX != null) && (BrakMX.Count > 0)),
                ((Roz != null) && (Roz.Count > 0)), ((Otstoy != null) && (Otstoy.Count > 0)), ((Akt != null) && (Akt.Count > 0)),
                ((SobNugdy != null) && (SobNugdy.Count > 0)), ((Podarky != null) && (Podarky.Count > 0)),
                ((Prokat != null) && (Prokat.Count > 0)), ((Ilin != null) && (Ilin.Count > 0)));
        }

        private string ZaprosAnalog(DateTime startDate, bool IsTA)
        {
            string q = "";
            if (IsTA)
                q = @"
                    SELECT
                        rg.SP408 As nomk,
                        rg.SP411 As Ostatok
                    FROM
                        RG405 As rg (NOLOCK)
                    WHERE
                        rg.PERIOD = @Period and rg.SP418 not in ('" + Sklad.Otstoy + "') and rg.SP4062 in " + OwnFirms.InCondition;
            else
            {
                q = @"
                    SELECT
                        TMP.nomk AS nomk,
                        SUM(TMP.ostN + TMP.pr - TMP.rh) As Ostatok
                    FROM (
                        SELECT
                            rg.SP408 As nomk,
                            rg.SP411 As ostN,
                            0 As pr,
                            0 As rh
                        FROM
                            RG405 As rg (NOLOCK)
                        WHERE
                            rg.PERIOD = @PeriodR and rg.SP418 not in ('" + Sklad.Otstoy + "') and rg.SP4062 in " + OwnFirms.InCondition;
                if (startDate != new DateTime(startDate.Year, startDate.Month, 1))
                    q = q + @"
                        UNION ALL
                        SELECT
                            ra1.SP408 As nomk,
                            (ra1.SP411*((ra1.DEBKRED+1)%2))- (ra1.SP411*ra1.DEBKRED) As ostN,
                            0 As pr,
                            0 As rh
                        FROM
                            RA405 As ra1 (NOLOCK)
                        INNER JOIN _1SJOURN As TabJ (NOLOCK)
                            ON (ra1.IDDOC = TabJ.IDDOC)
                        WHERE
                            TabJ.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN1, 112) AND TabJ.DATE_TIME_IDDOC < convert(varchar(8), @PeriodN, 112) and ra1.SP418 not in ('" + Sklad.Otstoy + "') and ra1.SP4062 in " + OwnFirms.InCondition;
                q = q + @"
                        UNION ALL
                        SELECT
                            ra2.SP408 As nomk,
                            0 As ostN,
                            (ra2.SP411*((ra2.DEBKRED+1)%2)) As pr,
                            (ra2.SP411*ra2.DEBKRED) As rh
                        FROM
                            RA405 As ra2 (NOLOCK)
                        INNER JOIN _1SJOURN As TabJ (NOLOCK)
                            ON (ra2.IDDOC = TabJ.IDDOC)
                        WHERE
                            TabJ.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112) AND TabJ.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112) and ra2.SP418 not in ('" + Sklad.Otstoy + @"') and ra2.SP4062 in " + OwnFirms.InCondition + @"
                        ) AS tmp GROUP BY nomk";
            }
            return q;
        }

        private DateTime GetDateTA()
        {
            DateTime TA = DateTime.Now;
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string querySystem = "select top 1 curdate from _1ssystem;";
                SqlCommand commandSystem = new SqlCommand(querySystem, connection);
                connection.Open();

                SqlDataReader reader = commandSystem.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                            TA = reader.GetDateTime(0);
                    }
                }
                catch
                {
                }
                finally
                {
                    reader.Close();
                }
            }

            return TA;
        }

        [WebMethod(Description = "Номенклатура, проданная контрагенту за период")]
        public NomenklaturaProdanaReply NomenklaturaProdana(string CustomerCode = null)
        {
            NomenklaturaProdanaReply reply = new NomenklaturaProdanaReply();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            List<string> listCustomer = new List<string>(CustomerCode.Split(','));
            if ((listCustomer.Count == 0) || ((listCustomer.Count == 1) && (listCustomer[0] == "")))
            {
                reply.ReplyState.Description = "Необходимо указать контрагента";
                return reply;
            }

            DateTime startDate = new DateTime(2016, 1, 1); //always 01.01.2016
            DateTime DateTA = GetDateTA();
            DateTime startMonth = new DateTime(DateTA.Year, DateTA.Month, 1);
            DateTime endDate = startMonth.AddMonths(-1);
            using (SqlConnection connection = new SqlConnection(connWebStr))
            {
                string queryString = @"
                        select cust_code, nomk_id, nomk_code, nomk_art, nomk_descr, 
	                         sum(nach_ost) as n_ost, sum(prih) as prihod, sum(ras) as rashod 
                        from
	                        (
	                        select cust.code as cust_code, nomk.id as nomk_id, nomk.code as nomk_code, nomk.sp85 as nomk_art, nomk.descr as nomk_descr, 
		                         rg.sp2375 as nach_ost, 0 as prih, 0 as ras
                                     from RG2351 as rg (NOLOCK)
                                     inner join sc84 as nomk (NOLOCK) on (nomk.id = rg.sp2343)
                                     inner join sc172 as cust (NOLOCK) on (cust.id = rg.sp2344)
                                     where 
	                                    (
			                         cust.code in ({0}) or
                                           cust.sp9631 in (select distinct sp9631 from sc172 (NOLOCK) where code in ({0}) and sp9631 <> '     0   ')
                                           ) 
			                         and rg.period between @startPeriod and @endPeriod
                                     and rg.SP4067 in ('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + @"')

	                        union all
	                        select cust.code as cust_code, nomk.id as nomk_id, nomk.code as nomk_code, nomk.sp85 as nomk_art, nomk.descr as nomk_descr,
		                         0 as nach_ost, (ra.sp2375*((DEBKRED+1)%2)) as prih, (ra.sp2375*DEBKRED) as ras
		                         from RA2351 as ra (NOLOCK)
    		                         inner join _1SJOURN as TabJ (NOLOCK) on (ra.IDDOC = TabJ.IDDOC)
                                     inner join sc84 as nomk (NOLOCK) on (nomk.id = ra.sp2343)
                                     inner join sc172 as cust (NOLOCK) on (cust.id = ra.sp2344)
		                         where
	                                    (
			                         cust.code in ({0}) or
                                           cust.sp9631 in (select distinct sp9631 from sc172 (NOLOCK) where code in ({0}) and sp9631 <> '     0   ')
                                           ) 
		    	                        and TabJ.DATE_TIME_IDDOC >= convert(varchar(8), @PeriodN, 112)
        	 	                        and TabJ.DATE_TIME_IDDOC < convert(varchar(8), @PeriodK, 112)
                                     and ra.SP4067 in ('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + @"')
	                        ) as TMP
                        group by cust_code, nomk_id, nomk_code, nomk_art, nomk_descr
                        order by cust_code, nomk_descr
                      ";
                queryString = string.Format(queryString, string.Join(",", listCustomer.Select(x => "'" + x + "'").ToArray()));
                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 2000;
                command.Parameters.Add("@startPeriod", SqlDbType.DateTime);
                command.Parameters.Add("@endPeriod", SqlDbType.DateTime);
                command.Parameters.Add("@PeriodN", SqlDbType.DateTime);
                command.Parameters.Add("@PeriodK", SqlDbType.DateTime);
                command.Parameters["@startPeriod"].Value = startDate;
                command.Parameters["@endPeriod"].Value = endDate;
                command.Parameters["@PeriodN"].Value = startMonth;
                command.Parameters["@PeriodK"].Value = DateTA.AddDays(1);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                reply.Tovars = new List<CustTovar>();
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            reply.Tovars.Add(new CustTovar { CustCode = reader.GetString(0).Trim(), Id = reader.GetString(1).Trim(), Code = reader.GetString(2).Trim(), Name = reader.GetString(4).Trim() });
                        }
                    }
                    reply.ReplyState.State = "Success";
                    reply.ReplyState.Description = "Информация успешно извлечена из БД";

                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }

            return reply;
        }

        private void BlockGarantii(SqlConnection connection, CompleteRepairStatus crs)
        {
            string queryDocuments = @"select МестоХранения.descr, left(j.date_time_iddoc,8) as ДатаПеремещения
                                            from ra11495 as ГарантийныеРемонты (NOLOCK)
                                            inner join _1sjourn as j (NOLOCK) on (j.iddoc = ГарантийныеРемонты.iddoc)
                                            inner join SC11473 as МестоХранения (NOLOCK) on (МестоХранения.id = ГарантийныеРемонты.SP11493)
                                            where j.date_time_iddoc >= @ДатаЗавершения and ГарантийныеРемонты.SP11491 = @ДокЗавершения and ГарантийныеРемонты.DEBKRED = 0
                                            order by j.date_time_iddoc desc
                                        ";
            string queryGarZaver = @"select top 1 ГарантийноеЗавершение.iddoc, left(j.date_time_iddoc,8) as ДатаГарОтчетности
                                        from RA12406 as ГарантийноеЗавершение (NOLOCK)
                                        inner join _1sjourn as j (NOLOCK) on (j.iddoc = ГарантийноеЗавершение.iddoc)
                                        where j.date_time_iddoc >= @ДатаЗавершения and ГарантийноеЗавершение.SP12400 = @ДокЗавершения and ГарантийноеЗавершение.DEBKRED = 1
                                        order by j.date_time_iddoc desc
                                    ";
            string queryGarOtch = @"select sum(ГарантийнаяОтчетность.SP11328) as Сумма
                                        from DT11330 as ГарантийнаяОтчетность (NOLOCK)
                                        where ГарантийнаяОтчетность.iddoc = @ДокГарОтчетность and ГарантийнаяОтчетность.SP11318 = @ДокЗавершения and 
                                        ГарантийнаяОтчетность.SP11370 = 0 and ГарантийнаяОтчетность.SP11369 < 2
                                   ";

            using (SqlCommand commandDocuments = new SqlCommand(queryDocuments, connection))
            {
                commandDocuments.Parameters.AddWithValue("@ДатаЗавершения", crs.DateZaver.ToString("yyyyMMdd"));
                commandDocuments.Parameters.AddWithValue("@ДокЗавершения", crs.DocZaverId);
                using (SqlDataReader reader = commandDocuments.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string MX = reader["descr"].ToString().Trim().ToUpper();
                        DateTime DateDoc = DateTime.ParseExact(reader["ДатаПеремещения"].ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
                        if (MX == "ДОСТАВКА")
                        {
                            crs.DatePer = DateDoc;
                        }
                        else if (MX == "КОММЕРЧЕСКИЙ ОТДЕЛ")
                        {
                            crs.DatePri = DateDoc;
                        }
                    }
                }
            }
            string GarOtchID = "";
            using (SqlCommand command = new SqlCommand(queryGarZaver, connection))
            {
                command.Parameters.AddWithValue("@ДатаЗавершения", crs.DateZaver.ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@ДокЗавершения", crs.DocZaverId);
                using (SqlDataReader reader = command.ExecuteReader(CommandBehavior.SingleRow))
                {
                    if (reader.Read())
                    {
                        GarOtchID = reader["iddoc"].ToString();
                        crs.DateGarOtch = DateTime.ParseExact(reader["ДатаГарОтчетности"].ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    }
                }
            }
            if (!string.IsNullOrEmpty(GarOtchID))
            {
                using (SqlCommand command = new SqlCommand(queryGarOtch, connection))
                {
                    command.Parameters.AddWithValue("@ДокГарОтчетность", GarOtchID);
                    command.Parameters.AddWithValue("@ДокЗавершения", crs.DocZaverId);
                    using (SqlDataReader reader = command.ExecuteReader(CommandBehavior.SingleRow))
                    {
                        if (reader.Read())
                            crs.SummaGarOtch = Math.Round(Convert.ToDecimal(reader["Сумма"], CultureInfo.InvariantCulture), 2);
                    }
                }
            }
        }

        [WebMethod(Description = "Консолидированный отчет мастерской")]
        public RepairStatusArray ConsolidatedServiceReport(DateTime StartDate, string RemontType = "", string RemontVid = "", string Brend = "",
            string Master = "", string Avtor = "", string Postav = "")
        {
            RepairStatusArray reply = new RepairStatusArray();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";
            DateTime DateTA = GetDateTA();

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string query = @"select tmp.НомерКвитанции, tmp.ДатаКвитанции
                                from
                                (
                                select sp9969 as НомерКвитанции, sp10084 as ДатаКвитанции
                                    from rg9972 (NOLOCK)
                                    where rg9972.period = @period

                                union all

                                select pm.sp9969 as НомерКвитанции, pm.sp10084 as ДатаКвитанции
                                    from ra9972 as pm (NOLOCK)
                                    inner join _1sjourn as j (NOLOCK) on (j.iddoc = pm.iddoc)
                                    where j.date_time_iddoc >= @periodN and j.date_time_iddoc <= @periodK
                                ) as tmp
                                group by tmp.НомерКвитанции, tmp.ДатаКвитанции
                                order by tmp.ДатаКвитанции, tmp.НомерКвитанции
                                ";
                string queryDocPriem = @"select top 1 docPriem.SP9891 as ЗавНомер, left(j.date_time_iddoc,8) as ДатаПриема,
                                            IsNull(spKontr.descr,'') as Контрагент, spNomk.descr as Изделие, 
                                            IsNull(spBrend.descr,'<не указан>') as Brend, IsNull(spBrend.id,'') as BrendID,
                                            IsNull(spPostav.id,'') as ПоставщикID, IsNull(spPostav.descr,'') as Поставщик, 
                                            spPolz.descr as Автор, spPolz.id as АвторID
                                            from ra10471 as raPriem (NOLOCK)
                                            inner join dh9899 as docPriem (NOLOCK) on (docPriem.iddoc = raPriem.SP10469)
                                            inner join _1sjourn as j (NOLOCK) on (j.iddoc = docPriem.iddoc)
                                            inner join sc30 as spPolz (NOLOCK) on (j.sp74 = spPolz.id)
                                            inner join sc84 as spNomk (NOLOCK) on (docPriem.SP9890 = spNomk.id)
                                            left join sc8840 as spBrend (NOLOCK) on (spNomk.sp8842 = spBrend.id)
                                            left join sc172 as spPostav (NOLOCK) on (spBrend.SP11193 = spPostav.id)
                                            left join sc172 as spKontr (NOLOCK) on (docPriem.SP9889 = spKontr.id)
                                            where raPriem.SP10467 = @НомерКвитанции and raPriem.SP10468 = @ДатаКвитанции
                                        ";
                string queryDocZaver = @"select top 1 docZaver.iddoc, left(j.date_time_iddoc,8) as ДатаЗавершения, docZaver.SP10455 as Сумма,
                                            docZaver.SP10449 as СуммаЗапчастей, docZaver.SP10450 as СуммаЗапчастейГар,
                                            docZaver.SP10453 as СуммаРабот, docZaver.SP10454 as СуммаРаботГар
                                            from ra10476 as raZaver (NOLOCK)
                                            inner join dh10457 as docZaver (NOLOCK) on (raZaver.SP10474 = docZaver.iddoc)
                                            inner join _1sjourn as j (NOLOCK) on (j.iddoc = docZaver.iddoc)
                                            where raZaver.SP10472 = @НомерКвитанции and raZaver.SP10473 = @ДатаКвитанции
                                        ";

                string queryPartii = @"select top 1 pm.SP9963 as Статус, pm.SP9958 as Гарантия, IsNull(spMastera.descr,'') as Мастер,
                                        IsNull(spMastera.id,'') as МастерID
                                        from rg9972 as pm (NOLOCK)
                                        left join RG11049 as rOstIzdel (NOLOCK) on ((pm.sp9969 = rOstIzdel.SP11042) and (pm.sp10084 = rOstIzdel.SP11043) and (rOstIzdel.SP11047 > 0) and (rOstIzdel.period = @period))
                                        left join sc9864 as spMastera (NOLOCK) on (rOstIzdel.SP11046 = spMastera.ID)
                                        where pm.period = @period and pm.SP9969 = @НомерКвитанции and pm.SP10084 = @ДатаКвитанции
                                      ";
                string queryRez = @"select top 1 spPolz.id as АвторID, spPolz.descr as Автор, docRez.SP11032 as Сумма, docRez.SP11027 as СуммаЗапчастей, docRez.SP11028 as СуммаЗапчастейГар,
                                        docRez.SP11030 as СуммаРабот, docRez.SP11031 as СуммаРаботГар
                                        from dh11037 as docRez (NOLOCK)
                                        inner join _1sjourn as j (NOLOCK) on (docRez.iddoc = j.iddoc)
                                        inner join sc30 as spPolz (NOLOCK) on (j.sp74 = spPolz.id)
                                        where docRez.sp11006 = @НомерКвитанции and docRez.sp11007 = @ДатаКвитанции and j.closed = 1
                                   ";
                string queryDocVyd = @"select top 1 left(j.date_time_iddoc,8) as ДатаВыдачи
                                        from DH10054 as docVyd (NOLOCK)
                                        inner join _1sjourn as j (NOLOCK) on (docVyd.iddoc = j.iddoc)
                                        where docVyd.SP10036 = @НомерКвитанции and docVyd.SP10037 = @ДатаКвитанции and
                                            j.closed = 1 and docVyd.SP10461 >= 0
                                      ";

                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.Add("@period", SqlDbType.DateTime);
                command.Parameters.Add("@periodN", SqlDbType.VarChar);
                command.Parameters.Add("@periodK", SqlDbType.VarChar);
                DateTime Period = new DateTime(StartDate.Year, StartDate.Month, 1);
                command.Parameters["@period"].Value = Period;
                command.Parameters["@periodN"].Value = StartDate.ToString("yyyyMMdd");
                command.Parameters["@periodK"].Value = DateTA.ToString("yyyyMMdd");

                SqlCommand commandDocPriem = new SqlCommand(queryDocPriem, connection);
                commandDocPriem.Parameters.Add("@НомерКвитанции", SqlDbType.VarChar);
                commandDocPriem.Parameters.Add("@ДатаКвитанции", SqlDbType.Int);

                SqlCommand commandDocZaver = new SqlCommand(queryDocZaver, connection);
                commandDocZaver.Parameters.Add("@НомерКвитанции", SqlDbType.VarChar);
                commandDocZaver.Parameters.Add("@ДатаКвитанции", SqlDbType.Int);

                SqlCommand commandPM = new SqlCommand(queryPartii, connection);
                commandPM.Parameters.AddWithValue("@period", Period);
                commandPM.Parameters.Add("@НомерКвитанции", SqlDbType.VarChar);
                commandPM.Parameters.Add("@ДатаКвитанции", SqlDbType.Int);

                SqlCommand commandDocRez = new SqlCommand(queryRez, connection);
                commandDocRez.Parameters.Add("@НомерКвитанции", SqlDbType.VarChar);
                commandDocRez.Parameters.Add("@ДатаКвитанции", SqlDbType.Int);

                SqlCommand commandDocVyd = new SqlCommand(queryDocVyd, connection);
                commandDocVyd.Parameters.Add("@НомерКвитанции", SqlDbType.VarChar);
                commandDocVyd.Parameters.Add("@ДатаКвитанции", SqlDbType.Int);

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                List<CompleteRepairStatus> rs = new List<CompleteRepairStatus>();
                try
                {
                    while (reader.Read())
                    {
                        CompleteRepairStatus crs = new CompleteRepairStatus();
                        crs.KvitNumber = reader["НомерКвитанции"].ToString().Trim();
                        crs.KvitDate = Convert.ToInt16(reader["ДатаКвитанции"]);

                        commandDocPriem.Parameters["@НомерКвитанции"].Value = crs.KvitNumber;
                        commandDocPriem.Parameters["@ДатаКвитанции"].Value = crs.KvitDate;
                        using (SqlDataReader readerDocPriem = commandDocPriem.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (readerDocPriem.Read())
                            {
                                crs.ZavNumber = readerDocPriem["ЗавНомер"].ToString().Trim();
                                crs.DatePriema = DateTime.ParseExact(readerDocPriem["ДатаПриема"].ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
                                crs.Kontragent = readerDocPriem["Контрагент"].ToString().Trim();
                                crs.Izdelie = readerDocPriem["Изделие"].ToString().Trim();
                                crs.Brend = readerDocPriem["Brend"].ToString().Trim();
                                crs.BrendID = readerDocPriem["BrendID"].ToString();
                                crs.Postav = readerDocPriem["Поставщик"].ToString().Trim();
                                crs.PostavID = readerDocPriem["ПоставщикID"].ToString();
                                crs.Avtor = readerDocPriem["Автор"].ToString().Trim();
                                crs.AvtorID = readerDocPriem["АвторID"].ToString();
                                crs.Manager = "";
                            }
                        }
                        if (DateTime.Compare(crs.DatePriema, DateTime.MinValue) <= 0)
                        {
                            crs.ZavNumber = "";
                            crs.Kontragent = "";
                            crs.Izdelie = "";
                            crs.Brend = "";
                            crs.BrendID = "";
                            crs.Postav = "";
                            crs.PostavID = "";
                            crs.Avtor = "";
                            crs.AvtorID = "";
                            crs.Manager = "";
                        }
                        if (!string.IsNullOrEmpty(Brend))
                        {
                            if (crs.BrendID != Brend)
                                continue;
                        }
                        if (!string.IsNullOrEmpty(Postav))
                        {
                            if (crs.PostavID != Postav)
                                continue;
                        }
                        commandPM.Parameters["@НомерКвитанции"].Value = crs.KvitNumber;
                        commandPM.Parameters["@ДатаКвитанции"].Value = crs.KvitDate;
                        using (SqlDataReader readerPM = commandPM.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (readerPM.Read())
                            {
                                crs.Status = GetStatus(readerPM["Статус"].ToString().Trim());
                                crs.RepairType = GetRepairType(Convert.ToInt16(readerPM["Гарантия"]));
                                crs.Master = readerPM["Мастер"].ToString().Trim();
                                crs.MasterID = readerPM["МастерID"].ToString();
                            }
                        }
                        if (string.IsNullOrEmpty(crs.RepairType))
                        {
                            crs.Status = "";
                            crs.RepairType = "";
                            crs.Master = "";
                            crs.MasterID = "";
                        }
                        if (!string.IsNullOrEmpty(Master))
                        {
                            if (crs.MasterID != Master)
                                continue;
                        }
                        if (!string.IsNullOrEmpty(RemontType))
                        {
                            switch (RemontType)
                            {
                                case "П":
                                    if (crs.RepairType != "Платный")
                                        continue;
                                    break;
                                case "Г":
                                    if (crs.RepairType != "Гарантийный")
                                        continue;
                                    break;
                                case "ПП":
                                    if (crs.RepairType != "Предпродажный")
                                        continue;
                                    break;
                                case "Э":
                                    if (crs.RepairType != "Экспертиза")
                                        continue;
                                    break;
                                case "ГП":
                                    if (!((crs.RepairType == "Гарантийный") || (crs.RepairType == "Предпродажный")))
                                        continue;
                                    break;
                            }
                        }
                        commandDocRez.Parameters["@НомерКвитанции"].Value = crs.KvitNumber;
                        commandDocRez.Parameters["@ДатаКвитанции"].Value = crs.KvitDate;
                        using (SqlDataReader readerRez = commandDocRez.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (readerRez.Read())
                            {
                                crs.Avtor = readerRez["Автор"].ToString().Trim();
                                crs.AvtorID = readerRez["АвторID"].ToString();
                                if ((crs.RepairType == "Платный") || ((crs.RepairType == "Экспертиза") && (crs.Status == "Экспертиза завершена")))
                                {
                                    crs.PredvarSumma = Math.Round(Convert.ToDecimal(readerRez["Сумма"], CultureInfo.InvariantCulture), 2);
                                    crs.PredvarSummaZ = Math.Round(Convert.ToDecimal(readerRez["СуммаЗапчастей"], CultureInfo.InvariantCulture), 2);
                                    crs.PredvarSummaR = Math.Round(Convert.ToDecimal(readerRez["СуммаРабот"], CultureInfo.InvariantCulture), 2);
                                }
                                else
                                {
                                    crs.PredvarSummaZ = Math.Round(Convert.ToDecimal(readerRez["СуммаЗапчастейГар"], CultureInfo.InvariantCulture), 2);
                                    crs.PredvarSummaR = Math.Round(Convert.ToDecimal(readerRez["СуммаРаботГар"], CultureInfo.InvariantCulture), 2);
                                    crs.PredvarSumma = crs.PredvarSummaZ + crs.PredvarSummaR;
                                }
                            }
                        }
                        if (!string.IsNullOrEmpty(Avtor))
                        {
                            if (crs.AvtorID != Avtor)
                                continue;
                        }
                        commandDocZaver.Parameters["@НомерКвитанции"].Value = crs.KvitNumber;
                        commandDocZaver.Parameters["@ДатаКвитанции"].Value = crs.KvitDate;
                        using (SqlDataReader readerZaver = commandDocZaver.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (readerZaver.Read())
                            {
                                crs.DocZaverId = readerZaver["iddoc"].ToString();
                                crs.DateZaver = DateTime.ParseExact(readerZaver["ДатаЗавершения"].ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
                                if (crs.RepairType == "Платный")
                                {
                                    crs.Summa = Math.Round(Convert.ToDecimal(readerZaver["Сумма"], CultureInfo.InvariantCulture), 2);
                                    crs.SummaZ = Math.Round(Convert.ToDecimal(readerZaver["СуммаЗапчастей"], CultureInfo.InvariantCulture), 2);
                                    crs.SummaR = Math.Round(Convert.ToDecimal(readerZaver["СуммаРабот"], CultureInfo.InvariantCulture), 2);
                                }
                                else
                                {
                                    crs.SummaZ = Math.Round(Convert.ToDecimal(readerZaver["СуммаЗапчастейГар"], CultureInfo.InvariantCulture), 2);
                                    crs.SummaR = Math.Round(Convert.ToDecimal(readerZaver["СуммаРаботГар"], CultureInfo.InvariantCulture), 2);
                                    crs.Summa = crs.SummaZ + crs.SummaR;
                                }
                            }
                        }
                        commandDocVyd.Parameters["@НомерКвитанции"].Value = crs.KvitNumber;
                        commandDocVyd.Parameters["@ДатаКвитанции"].Value = crs.KvitDate;
                        using (SqlDataReader readerVyd = commandDocVyd.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (readerVyd.Read())
                            {
                                crs.DateVyd = DateTime.ParseExact(readerVyd["ДатаВыдачи"].ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
                            }
                        }

                        if ((crs.RepairType != "Платный") && (!string.IsNullOrEmpty(crs.DocZaverId)))
                        {
                            BlockGarantii(connection, crs);
                        }
                        if (!string.IsNullOrEmpty(RemontVid))
                        {
                            switch (RemontVid)
                            {
                                case "НЕ ЗАВЕРШЕН":
                                    if (crs.DateZaver > DateTime.MinValue)
                                        continue;
                                    break;
                                case "ЗАВЕРШЕН":
                                    if (crs.DateZaver == DateTime.MinValue)
                                        continue;
                                    break;
                                case "НЕ ВЫДАН":
                                    if (crs.DateVyd > DateTime.MinValue)
                                        continue;
                                    break;
                                case "НЕ ПЕРЕДАН":
                                    if (!((crs.RepairType == "Гарантийный") || (crs.RepairType == "Предпродажный")))
                                        continue;
                                    if ((crs.DateZaver == DateTime.MinValue) || (crs.DatePer > DateTime.MinValue) || (crs.DatePri > DateTime.MinValue))
                                        continue;
                                    break;
                                case "НЕ ПРИНЯТ":
                                    if (!((crs.RepairType == "Гарантийный") || (crs.RepairType == "Предпродажный")))
                                        continue;
                                    if ((crs.DateZaver == DateTime.MinValue) || ((crs.DatePer == DateTime.MinValue) && (crs.DatePri == DateTime.MinValue)) || ((crs.DatePer == DateTime.MinValue) && (crs.DatePri > DateTime.MinValue)) || ((crs.DatePer > DateTime.MinValue) && (crs.DatePri > DateTime.MinValue)))
                                        continue;
                                    break;
                                case "НЕ ОТЧИТАН":
                                    if (!((crs.RepairType == "Гарантийный") || (crs.RepairType == "Предпродажный")))
                                        continue;
                                    if ((crs.DatePri == DateTime.MinValue) || (crs.DateGarOtch > DateTime.MinValue))
                                        continue;
                                    break;
                            }
                        }
                        rs.Add(crs);
                    }

                    reply.RepairStatuses = rs.ToArray();

                    reply.KolOpenRemont = rs.Count(v => v.DateZaver == DateTime.MinValue);
                    reply.KolClosedRemont = rs.Count(v => v.DateZaver > DateTime.MinValue);
                    reply.KolVydRemont = rs.Count(v => v.DateVyd > DateTime.MinValue);
                    reply.KolNeVydRemont = rs.Count(v => v.DateVyd == DateTime.MinValue);
                    reply.Summa = rs.Sum(crs => crs.DateZaver == DateTime.MinValue ? crs.PredvarSumma : crs.Summa);
                    reply.SummaZ = rs.Sum(crs => crs.DateZaver == DateTime.MinValue ? crs.PredvarSummaZ : crs.SummaZ);
                    reply.SummaR = rs.Sum(crs => crs.DateZaver == DateTime.MinValue ? crs.PredvarSummaR : crs.SummaR);
                    reply.SrTimeOpenRemont = Math.Round(rs.Select(crs => ((crs.DateZaver == DateTime.MinValue ? DateTime.Today : crs.DateZaver) - crs.DatePriema).TotalDays).DefaultIfEmpty(0).Average());
                    reply.SrTimeClosedRemont = Math.Round(rs.Where(x => x.DateZaver > DateTime.MinValue).Select(crs => ((crs.DateVyd == DateTime.MinValue ? DateTime.Today : crs.DateVyd) - crs.DateZaver).TotalDays).DefaultIfEmpty(0).Average());
                    reply.SrTimeVydRemont = Math.Round(rs.Where(x => x.DateVyd > DateTime.MinValue).Select(crs => (crs.DateVyd - crs.DateZaver).TotalDays).DefaultIfEmpty(0).Average());
                    reply.SrSummaOneRemont = Math.Round(reply.Summa / (reply.KolOpenRemont + reply.KolClosedRemont));

                    reply.KolPerDocument = rs.Count(x => x.DatePer > DateTime.MinValue);
                    reply.KolPriDocument = rs.Count(x => x.DatePri > DateTime.MinValue);
                    reply.SummaPerDocument = rs.Where(x => x.DatePer > DateTime.MinValue).Select(crs => crs.Summa).DefaultIfEmpty(0).Sum();
                    reply.SummaPriDocument = rs.Where(x => x.DatePri > DateTime.MinValue).Select(crs => crs.Summa).DefaultIfEmpty(0).Sum();

                    reply.KolOtchOtp = rs.Count(x => x.DateGarOtch > DateTime.MinValue);
                    reply.SummaOtchOtp = rs.Where(x => x.DateGarOtch > DateTime.MinValue).Select(y => y.SummaGarOtch).DefaultIfEmpty(0).Sum();
                    reply.Delta = reply.SummaOtchOtp - rs.Where(x => x.DateGarOtch > DateTime.MinValue).Select(y => y.Summa).DefaultIfEmpty(0).Sum();

                    reply.ReplyState.State = "Success";
                    reply.ReplyState.Description = "Информация успешно извлечена из БД";

                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                    command.Dispose();
                    commandDocPriem.Dispose();
                    commandDocRez.Dispose();
                    commandDocVyd.Dispose();
                    commandDocZaver.Dispose();
                    commandPM.Dispose();
                    connection.Close();
                }
            }

            return reply;
        }

        [WebMethod(Description = "Создание Excel по шаблону")]
        public PriceList CreateExcelByTemplate(string template, List<nDataEntry> Data)
        {
            PriceList Response = new PriceList();
            Response.ReplyState = new ReplyState();
            Response.ReplyState.State = "Error";
            Response.ReplyState.Description = "Ошибка";
            try
            {
                Response.ExcelFile = FileXLS.CreateExcelFromTemplate(template, Data);

                Response.ReplyState.State = "Success";
                Response.ReplyState.Description = "Файл успешно создан";
            }
            catch (Exception e)
            {
                Response.ReplyState.State = "Error";
                Response.ReplyState.Description = e.Message;
                var st = new System.Diagnostics.StackTrace(e, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                Response.ReplyState.Description += Environment.NewLine + frame.GetFileLineNumber();
            }
            return Response;
        }
        [WebMethod(Description = "Загрузка по Яндексу метод FBY")]
        public StandartReply LoadYandexFBY(byte[] yandexFile, string yandexFileExtension, byte[] reportFile, string reportFileExtension, string sheetName)
        {
            StandartReply reply = new StandartReply();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            if (string.IsNullOrEmpty(yandexFileExtension))
                yandexFileExtension = "xlsx";
            if (string.IsNullOrEmpty(reportFileExtension))
                reportFileExtension = "xlsx";

            try
            {
                List<RowYandexFBY> report = FileXLS.LoadReportExcel(reportFileExtension, reportFile, sheetName);

                using (SqlConnection connection = new SqlConnection(connStr))
                {
                    string queryString = @"
                        select id, IsNull(rozn,0) as rozn, IsNull(rozn_sp,0) as rozn_sp 
                        from vzTovar
                        where id in {0}
                        ";
                    queryString = string.Format(queryString, "(" + string.Join(", ", report.Select(x => "'" + x.Id + "'")) + ")");
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.CommandTimeout = 300;
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            double rozn = Convert.ToDouble(reader["rozn"], CultureInfo.InvariantCulture);
                            double rozn_sp = Convert.ToDouble(reader["rozn_sp"], CultureInfo.InvariantCulture);
                            double price = Math.Max(rozn, 0);
                            if (rozn_sp > 0)
                                price = Math.Min(rozn, rozn_sp);
                            foreach (var entry in report.Where(x => x.Id == (string)reader["id"]))
                            {
                                entry.Цена = price;
                            }
                        }
                    }
                }

                reply.ExcelFile = FileXLS.CreateResultYandexFBY(yandexFileExtension, yandexFile, report);

                reply.ReplyState.State = "Success";
                reply.ReplyState.Description = "Информация успешно извлечена из БД";
            }
            catch (Exception e)
            {
                reply.ReplyState.State = "Error";
                reply.ReplyState.Description = e.Message;
            }

            return reply;
        }
        private void ParseFilePath(string path, string defaultName, out string fileName, out string extension)
        {
            try
            {
                fileName = Path.GetFileNameWithoutExtension(path);
                extension = Path.GetExtension(path).Replace(".", "");
            }
            catch
            {
                fileName = defaultName;
                extension = "xlsx";
            }
        }
        private string ПривестиНомерКДлине(string Номер, int Длина)
        {
            string Префикс = "";
            foreach (var symbol in Номер)
            {
                if (char.IsLetter(symbol))
                {
                    Префикс += symbol;
                }
                else
                    break;
            }
            string ЧисловаяЧасть = Номер.Substring(Префикс.Length).PadLeft(Длина - Префикс.Length, '0');
            return Префикс + ЧисловаяЧасть;
        }
        [WebMethod(Description = "Отчет по остаткам товара фирмы")]
        public StandartReply ReportFirmaStock(string fileExtension, string firmaId, DateTime period)
        {
            StandartReply reply = new StandartReply();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            List<NomenklaturaFirma> table = new List<NomenklaturaFirma>();
            string firmaName = "";
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = @"select descr from sc4014 where id = @firmaId;
                                        select nomenk.id, nomenk.code, nomenk.descr, IsNull(nomenk.sp101,'') as polnDescr,
                                        IsNull(OKEI.code,'796') as codeEd, IsNull(OKEI.descr,'шт') as descrEd, 
                                        IsNull(nomenk.SP85,'') as article,
                                        IsNull(R_Part.ostatok, 0) as ostatok, IsNull(R_Part.summaUpr, 0) as sebest,
                                        IsNull(R_Part.ostatokP, 0) as ostatokP, IsNull(R_Part.summaP, 0) as sebestP,
										case 
											when nomenk.sp2417 = '   1X3   ' then 'услуга'
											else 'товар'
										end as nomenkType,
										case
											when nomenk.sp103 = '    I7   ' then '20%'
											when nomenk.sp103 = '    I9   ' then 'Без НДС'
											when nomenk.sp103 = '    I8   ' then '10%'
											when nomenk.sp103 = '   6F2   ' then '18%'
											when nomenk.sp103 = '   9YV   ' then '0%'
											else '20%'
										end as stavkaNDS,
										case
											when IsNull(country.SP11996,'') <> '' and IsNull(country.SP11996,'') <> '   ' then IsNull(country.SP11996,'')
											when IsNull(country.code,'') = '' then '643'
											else IsNull(country.code,'')
										end as countryCode,
										case when IsNull(country.descr,'') = '' then 'Россия' else IsNull(country.descr,'') end as countryDescr
                    from (select r.sp331, Sum(r.ostatok) as ostatok, Sum(r.ostatokP) as ostatokP, Sum(r.summaUpr) as summaUpr, Sum(r.summaP) as summaP from (
                        select SP331, Sum(SP342) as ostatok, 0 as ostatokP, Sum(SP421) as summaUpr, 0 as summaP from RG328
                        where period = @period and SP340 = '   2TB   ' and SP4061 = @firmaId
                        group by SP331 having Sum(SP342) <> 0 or Sum(SP421) <> 0
                        union all
                        select SP331, 0 as ostatok, Sum(SP342) as ostatokP, 0 as summaUpr, Sum(SP421) as summaP from RG328
                        where period = @period and SP340 = '   2TC   ' and SP4061 = @firmaId
                        group by SP331 having Sum(SP342) <> 0 or Sum(SP421) <> 0) as r group by r.sp331
                        ) as R_Part
                    inner join sc84 as nomenk on R_Part.sp331 = nomenk.id
                    left join SC75 as SpEd on (SpEd.id = nomenk.SP94)
                    left join sc41 as OKEI on SpEd.sp79 = OKEI.id
					LEFT JOIN _1SCONST TabConst (NOLOCK) ON nomenk.ID = TabConst.OBJID AND TabConst.ID = 5012 AND 
                                            TabConst.DATE = (SELECT MAX(TabConstl.DATE) FROM _1SCONST AS TabConstl (NOLOCK)
                                                                WHERE TabConstl.OBJID = TabConst.OBJID AND TabConstl.ID = TabConst.ID)
					left join sc566 as country on TabConst.VALUE = country.ID
                ";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 300;
                command.Parameters.AddWithValue("@period", new DateTime(period.Year, period.Month, 1).ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@firmaId", firmaId);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        firmaName = ((string)reader["descr"]).Trim();
                    }
                    if (reader.NextResult())
                        while (reader.Read())
                        {
                            table.Add(new NomenklaturaFirma
                            {
                                Код = ПривестиНомерКДлине((string)reader["code"], 11),
                                Наименование = ((string)reader["descr"]).Trim(),
                                ПолнНаименование = ((string)reader["polnDescr"]).Trim(),
                                Артикул = ((string)reader["article"]).Trim(),
                                ВидНоменклатуры = (string)reader["nomenkType"],
                                СтавкаНДС = (string)reader["stavkaNDS"],
                                ЕдиницаКод = (string)reader["codeEd"],
                                ЕдиницаНаименование = (string)reader["descrEd"],
                                СтранаКод = (string)reader["countryCode"],
                                СтранаНаименование = ((string)reader["countryDescr"]).Trim(),
                                Остаток = Convert.ToDouble(reader["ostatok"], CultureInfo.InvariantCulture),
                                Себестоимость = Convert.ToDouble(reader["sebest"], CultureInfo.InvariantCulture),
                                ОстатокПринятый = Convert.ToDouble(reader["ostatokP"], CultureInfo.InvariantCulture),
                                СебестоимостьПринятый = Convert.ToDouble(reader["sebestP"], CultureInfo.InvariantCulture),
                            });
                    }
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }
            //var fg = table.Where(x => x.Остаток < 0 || x.Себестоимость < 0);
            //var fы = table.Where(x => x.ОстатокПринятый < 0 || x.СебестоимостьПринятый < 0);

            //var rt = table.Where(x => x.Остаток == 0 && x.Себестоимость != 0);
            //var kj = table.Where(x => x.ОстатокПринятый == 0 && x.СебестоимостьПринятый != 0);

            //var qrt = table.Where(x => x.Остаток != 0 && x.ОстатокПринятый != 0);
            //var rrrr = table.GroupBy(x => x.Код).Where(x => x.Count() > 1);
            if (period.Day == DateTime.DaysInMonth(period.Year, period.Month))
                period = period.AddDays(1);

            reply.ExcelFile = FileXLS.CreateExcelОстаткиФирмы(table, firmaName, period, fileExtension);

            reply.ReplyState.State = "Success";
            reply.ReplyState.Description = "Информация успешно извлечена из БД";

            return reply;
        }
        private List<RowYandexStock> GetYandexStockValues(byte[] file, string path, byte[] stocksFile, string pathStocksFile)
        {
            ParseFilePath(pathStocksFile, "Stock", out string stockFileName, out string stockFileExtension);
            ParseFilePath(path, "File", out string fileName, out string fileExtension);

            var StockValues = FileXLS.GetSkuValuesfromExcel(stockFileExtension, stocksFile, 0, 5, 5);
            var Skus = FileXLS.GetSKUfromExcel(fileExtension, file, 0, 1);
            var SkuNotAllowed = StockValues.Where(x => !Skus.Any(y => y == x.Key)).Select(x => x.Key);
            var result = new List<RowYandexStock>();
            foreach (var ent in StockValues)
            {
                result.Add(new RowYandexStock { SKU = ent.Key, Allowed = !SkuNotAllowed.Contains(ent.Key), Quantity = ent.Value });
            }
            foreach (var sku in Skus.Where(x => !result.Any(y => y.SKU == x)))
            {
                result.Add(new RowYandexStock { SKU = sku, Allowed = true, Quantity = 0 });
            }
            return result;
        }
        [WebMethod(Description = "Отчет по Яндексу метод FBY")]
        public StandartReply ReportYandexFBY(
            byte[] file1, string path1, 
            byte[] stocksFile1, string pathStocksFile1,
            //byte[] file2, string path2,
            //byte[] stocksFile2, string pathStocksFile2,
            //byte[] file3, string path3,
            //byte[] stocksFile3, string pathStocksFile3,
            byte[] fileYandexPrice, string pathFileYandexPrice,
            byte[] fileFirstPrice, string pathFileFirstPrice,
            DateTime startDate, DateTime endDate)
        {
            StandartReply reply = new StandartReply();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            ParseFilePath(pathFileYandexPrice, "File", out string fileYandexPriceName, out string fileYandexPriceExtension);
            ParseFilePath(pathFileFirstPrice, "File", out string fileFirstPriceName, out string fileFirstPriceExtension);

            List<RowYandexFBY> table = new List<RowYandexFBY>();
            var YandexPriceValues = FileXLS.GetSkuValuesfromExcel(fileYandexPriceExtension, fileYandexPrice, 2, 26, 4, 1);
            var SKUs = YandexPriceValues.Select(x => x.Key).Distinct();

            var StockValues1 = GetYandexStockValues(file1, path1, stocksFile1, pathStocksFile1);
            //var StockValues2 = GetYandexStockValues(file2, path2, stocksFile2, pathStocksFile2);
            //var StockValues3 = GetYandexStockValues(file3, path3, stocksFile3, pathStocksFile3);
            var UnusedSKUs = SKUs.Where(x => !StockValues1.Any(y => y.SKU == x)).ToList();
            //var UnusedSKUs = SKUs.Where(x => !StockValues1.Any(y => y.SKU == x) && !StockValues2.Any(y => y.SKU == x) && !StockValues3.Any(y => y.SKU == x)).ToList();

            var FirstPriceValues = FileXLS.GetSkuValuesfromExcel(fileFirstPriceExtension, fileFirstPrice, 0, 7, 1);

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = @"
                    SELECT
                        sprNom.id As НоменклатураId, sprNom.code, sprNom.sp85 as Артикул, sprNom.Descr as Наименование, sprNom.sp14188 as Квант,
                        (IsNull(rg405.Остаток,0) - IsNull(rg4480.Резерв,0) - IsNull(rg11055.ОстатокАвСписка,0)) as СвободныйОстаток,
                        IsNull(rg464.ОжидаемыйПриход,0) as ОжидаемыйПриход,
                        IsNull(ra13979.Поставлено,0) as Поставлено,
                        IsNull(raP13979.Впути,0) as Впути,
                        IsNull(vzTovar.rozn,0) as rozn, IsNull(vzTovar.rozn_sp,0) as rozn_sp, IsNull(vzTovar.zakup,0) as zakup
                    FROM sc84 as sprNom
                    left join vzTovar on sprNom.id = vzTovar.id
                    left join (
                        select rg.sp408 As НоменклатураId, 
                            sum(rg.sp411) As Остаток
                        from rg405 As rg
                        where
                            rg.period = @tekRG and rg.sp418 = @skladId
                        group by rg.sp408) rg405 on sprNom.id = rg405.НоменклатураId
                    left join (
                        select rg.sp4477 As НоменклатураId, 
                            sum(rg.sp4479) As Резерв
                        from rg4480 As rg
                        where
                            rg.period = @tekRG and rg.sp4476 = @skladId
                        group by rg.sp4477) rg4480 on sprNom.id = rg4480.НоменклатураId
                    left join (
                        select rg.sp11050 As НоменклатураId, 
                            sum(rg.sp11054) As ОстатокАвСписка
                        from rg11055 As rg
                        where
                            rg.period = @tekRG and rg.sp11051 = @skladId
                        group by rg.sp11050) rg11055 on sprNom.id = rg11055.НоменклатураId
                    left join (
                        select rg.sp466 As НоменклатураId, 
                            sum(rg.sp4471) As ОжидаемыйПриход
                        from rg464 As rg
                        where
                            rg.period = @tekRG and rg.sp13166 <> 2
                        group by rg.sp466) rg464 on sprNom.id = rg464.НоменклатураId
                    left join (
                        select ra.SP13976 As НоменклатураId, 
                            sum(ra.SP13977*((ra.DEBKRED+1)%2)) As Поставлено
                        from RA13979 As ra
                        inner join _1SJOURN As j ON ra.IDDOC = j.IDDOC
                        inner join sc204 as d on ra.sp13974 = d.parentext
                        where
                            j.DATE_TIME_IDDOC >= @periodN AND j.DATE_TIME_IDDOC < @periodK and 
                            ra.sp13973 = @firmaId and (d.id = @dogovorId or d.id = @dogovorId2) 
                        group by ra.sp13976) ra13979 on sprNom.id = ra13979.НоменклатураId
                    left join (
                        select ra.SP13976 As НоменклатураId, 
                            sum(ra.SP13977*((ra.DEBKRED+1)%2)) As Впути
                        from RA13979 As ra
                        inner join _1SJOURN As j ON ra.IDDOC = j.IDDOC
                        inner join sc204 as d on ra.sp13974 = d.parentext
                        where
                            j.DATE_TIME_IDDOC >= @periodK AND j.DATE_TIME_IDDOC < @periodEnd and 
                            ra.sp13973 = @firmaId and (d.id = @dogovorId or d.id = @dogovorId2)
                        group by ra.sp13976) raP13979 on sprNom.id = raP13979.НоменклатураId
                    where
                        sprNom.code in {0}";
                queryString = string.Format(queryString, "(" + string.Join(", ", SKUs.Select(x => "'" + x + "'")) + ")");
                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 300;
                command.Parameters.AddWithValue("@tekRG", new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@periodN", startDate.ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@periodK", endDate.AddDays(1).ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@periodEnd", DateTime.Today.AddDays(1).ToString("yyyyMMdd"));
                command.Parameters.AddWithValue("@firmaId", Firma.IP_pavlov);
                command.Parameters.AddWithValue("@dogovorId", "   RLID  "); //Яндекс ООО
                command.Parameters.AddWithValue("@dogovorId2", "   IWXD  "); //Яндекс Маркет
                command.Parameters.AddWithValue("@skladId", Sklad.Ekran);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        double rozn = Convert.ToDouble(reader["rozn"], CultureInfo.InvariantCulture);
                        double rozn_sp = Convert.ToDouble(reader["rozn_sp"], CultureInfo.InvariantCulture);
                        double price = Math.Max(rozn, 0);
                        if (rozn_sp > 0)
                            price = Math.Min(rozn, rozn_sp);
                        var stock1 = StockValues1.FirstOrDefault(x => x.SKU == (string)reader["code"]);
                        //var stock2 = StockValues2.FirstOrDefault(x => x.SKU == (string)reader["code"]);
                        //var stock3 = StockValues3.FirstOrDefault(x => x.SKU == (string)reader["code"]);
                        table.Add(new RowYandexFBY
                        {
                            Id = (string)reader["НоменклатураId"],
                            Sku = (string)reader["code"],
                            Артикул = ((string)reader["Артикул"]).Trim(),
                            Наименование = ((string)reader["Наименование"]).Trim(),
                            Квант = Convert.ToDouble(reader["Квант"], CultureInfo.InvariantCulture),
                            СвободныйОстаток = Math.Max(Convert.ToDouble(reader["СвободныйОстаток"], CultureInfo.InvariantCulture), 0),
                            ОжидаемыйПриход = Convert.ToDouble(reader["ОжидаемыйПриход"], CultureInfo.InvariantCulture),
                            Поставлено = Convert.ToDouble(reader["Поставлено"], CultureInfo.InvariantCulture),
                            Впути = Convert.ToDouble(reader["Впути"], CultureInfo.InvariantCulture),
                            Цена = Convert.ToDouble(reader["rozn"], CultureInfo.InvariantCulture),
                            ЦенаСП = Convert.ToDouble(reader["rozn_sp"], CultureInfo.InvariantCulture),
                            ЦенаЗакуп = Convert.ToDouble(reader["zakup"], CultureInfo.InvariantCulture),
                            ЦенаНаЯндексе = YandexPriceValues.Where(x => x.Key == (string)reader["code"]).Select(x => x.Value).FirstOrDefault(),
                            ЦенаПервая = FirstPriceValues.Where(x => x.Key == (string)reader["code"]).Select(x => x.Value).FirstOrDefault(),
                            ОстаткиНаСкладеЯндекс1 = stock1 != null ? stock1.Quantity : 0,
                            Allowed1 = stock1 != null ? stock1.Allowed : false,
                            //ОстаткиНаСкладеЯндекс2 = stock2 != null ? stock2.Quantity : 0,
                            //Allowed2 = stock2 != null ? stock2.Allowed : false,
                            //ОстаткиНаСкладеЯндекс3 = stock3 != null ? stock3.Quantity : 0,
                            //Allowed3 = stock3 != null ? stock3.Allowed : false,
                        });
                    }
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }
            reply.ExcelFile = FileXLS.CreateExcelYandexFBY(table, startDate, endDate, SKUs.ToList(), "Отчет FBY", UnusedSKUs, "xlsx");

            reply.ReplyState.State = "Success";
            reply.ReplyState.Description = "Информация успешно извлечена из БД";

            return reply;
        }

        [WebMethod(Description = "Отчет Заказ по регламенту")]
        public ZakazReglamentReply ReportZakazReglament(int IsMaster = 0, string Extension = "xls", string StartDateStr = "", string EndDateStr = "", int fullInterf = 1,
                                                            int CostZak = 1, int PrihRas = 1, int Filial = 1, int Upak = 0, int MinOtgr = 0, int Sertif = 0, int flHarak = 0, int flBrend = 0,
                                                            string NomenkSelect = "", int NomenkInclude = 1,
                                                            string ProizSelect = "", int ProizInclude = 1, string FilialStr = "", string PpBrakStr = "", string BrakStr = "",
                                                            string RozStr = "", string OtstoyStr = "", string AktStr = "", string SobNugdyStr = "", string PodarkyStr = "", string ProkatStr = "", string IlinStr = "")
        {
            ZakazReglamentReply reply = new ZakazReglamentReply();
            reply.ReplyState = new ReplyState();
            reply.ReplyState.State = "Error";
            reply.ReplyState.Description = "Нет информации в БД";

            DateTime startDate;
            if (String.IsNullOrEmpty(StartDateStr))
                startDate = DateTime.Today;
            else
                startDate = DateTime.ParseExact(StartDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            DateTime endDate;
            bool IsTA = false;
            if (String.IsNullOrEmpty(EndDateStr))
            {
                endDate = DateTime.Today;
                IsTA = true;
            }
            else
            {
                if ((EndDateStr.Length > 10) && (EndDateStr.Substring(10) == "ТА"))
                {
                    IsTA = true;
                    endDate = DateTime.ParseExact(EndDateStr.Substring(0, 10), "dd.MM.yyyy", CultureInfo.InvariantCulture);
                }
                else
                    endDate = DateTime.ParseExact(EndDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            }
            bool FullInterface = (fullInterf == 1);
            bool S_CostZak = (CostZak == 1);
            bool S_PrihRas = (PrihRas == 1);
            bool S_Filial = (Filial == 1);
            bool S_Upak = (Upak == 1);
            bool S_MinOtgr = (MinOtgr == 1);
            bool S_Sertif = (Sertif == 1);

            List<string> listNomk;
            if (String.IsNullOrEmpty(NomenkSelect))
                listNomk = new List<string>();
            else
                listNomk = new List<string>(NomenkSelect.Split(','));
            if ((listNomk.Count == 1) && (listNomk[0].Trim() == ""))
                listNomk.Clear();
            List<string> listProiz;
            if (String.IsNullOrEmpty(ProizSelect))
                listProiz = new List<string>();
            else
                listProiz = new List<string>(ProizSelect.Split(','));
            if ((listProiz.Count == 1) && (listProiz[0].Trim() == ""))
                listProiz.Clear();
            List<string> listSkl;
            if (String.IsNullOrEmpty(FilialStr))
                listSkl = new List<string>();
            else
                listSkl = new List<string>(FilialStr.Split(','));
            List<string> PpBrak;
            if (String.IsNullOrEmpty(PpBrakStr))
                PpBrak = new List<string>();
            else
                PpBrak = new List<string>(PpBrakStr.Split(','));
            List<string> Brak;
            if (String.IsNullOrEmpty(BrakStr))
                Brak = new List<string>();
            else
                Brak = new List<string>(BrakStr.Split(','));
            List<string> Roz;
            if (String.IsNullOrEmpty(RozStr))
                Roz = new List<string>();
            else
                Roz = new List<string>(RozStr.Split(','));
            List<string> Otstoy;
            if (String.IsNullOrEmpty(OtstoyStr))
                Otstoy = new List<string>();
            else
                Otstoy = new List<string>(OtstoyStr.Split(','));
            List<string> Akt;
            if (String.IsNullOrEmpty(AktStr))
                Akt = new List<string>();
            else
                Akt = new List<string>(AktStr.Split(','));
            List<string> SobNugdy;
            if (String.IsNullOrEmpty(SobNugdyStr))
                SobNugdy = new List<string>();
            else
                SobNugdy = new List<string>(SobNugdyStr.Split(','));
            List<string> Podarky;
            if (String.IsNullOrEmpty(PodarkyStr))
                Podarky = new List<string>();
            else
                Podarky = new List<string>(PodarkyStr.Split(','));
            List<string> Prokat;
            if (String.IsNullOrEmpty(ProkatStr))
                Prokat = new List<string>();
            else
                Prokat = new List<string>(ProkatStr.Split(','));
            List<string> Ilin;
            if (String.IsNullOrEmpty(IlinStr))
                Ilin = new List<string>();
            else
                Ilin = new List<string>(IlinStr.Split(','));
            List<SkladProperties> SklProperties = new List<SkladProperties>();
            foreach (string Sk in listSkl)
            {
                SkladProperties sp = new SkladProperties();
                sp.Code = Sk;
                SklProperties.Add(sp);
            }

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                string queryString = "select id, code, descr from sc55 (NOLOCK) where code in (" + string.Join(",", listSkl.Select(x => "'" + x + "'").ToArray()) + ");";
                queryString = queryString + "select m.id as Id, m.DESCR as Name, m.sp14156 as ShortName, m.PARENTEXT as Parent, m.SP14155 as MarketType, m.SP14164 as Model from sc14042 m (NOLOCK) where (m.ISMARK = 0) order by m.SP14157;";
                queryString = queryString + @"select mu.parentext as nomId, mu.sp14147 as marketId, mu.sp14198 as comission from sc14152 mu (NOLOCK) 
                                        inner join sc84 as SpNom (NOLOCK) on (SpNom.id = mu.parentext)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842
                                        where (mu.ISMARK = 0) " +
                                            ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                            ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                            ";";
                queryString = queryString + "select R.nomk, SpEd.SP78, R.regTZ, R.Ostatok " +
                                            ZaprosTZ(startDate, IsTA) +
                                        @"inner join sc84 as SpNom (NOLOCK) on (SpNom.id = R.nomk) 
                                        inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842    
                                        where (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                              //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                            ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                            ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                            ";";
                queryString = queryString + @"select R.nomk, SpEd.SP78, R.sklad, (R.NachOst + R.NabN), R.Prihod, R.Rashod, (R.KonOst + R.NabK), R.NabN, R.NabK, R.RezN, R.RezK " +
                                            ZaprosOstatkiSklad(startDate, endDate, IsTA) +
                                        @"inner join sc84 as SpNom (NOLOCK) on (SpNom.id = R.nomk) 
                                        inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842    
                                        where (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                              //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                            ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                            ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                            ";";
                queryString = queryString + "select R.nomk, SpEd.SP78, R.sklad, convert(DateTime, R.MinDocDate, 112) as MinDocDT, convert(DateTime, R.MaxDocDate, 112) as MaxDocDT, R.prodano, R.prStoimost, R.sebest, R.PoGarantii, R.PoGarantiiSt, R.PoGarantiiSebeSt, R.Platno, R.PlatnoSt, R.PlatnoSebeSt " +
                                            FixZaprosProdag +
                                        @"inner join sc84 as SpNom (NOLOCK) on (SpNom.id = R.nomk) 
                                        inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842 
                                        where (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                              //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                            ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                            ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                            ";";
                queryString = queryString + @"select SpNom.id as nomk, SpEd.SP78 as koff, SpAnalog.SP11395 as analog, IsNull(R.Ostatok,0) as ostatok 
                                                from SC552 as SpAnalog (NOLOCK)
                                                left join (" + ZaprosAnalog(startDate, IsTA) + @") as R on (R.nomk = SpAnalog.SP11395)
                                                inner join sc84 as SpNom (NOLOCK) on (SpNom.id = SpAnalog.parentext) 
                                                inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                                inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842 
                                                where (SpAnalog.IsMark = 0) and (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                                                                //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                                    ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                                    ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                                    ";";
                queryString = queryString + @"select SpNom.id as nomk, SpNalichF.SP9791 as sklad, IsNull(TabConstNF.Value, 0) as nalFil, SpNom.SP10535 as koff 
                                                from SC9795 as SpNalichF (NOLOCK)
                                                inner join sc84 as SpNom (NOLOCK) on (SpNom.id = SpNalichF.parentext) 
                                                inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842 
                                                LEFT OUTER JOIN _1SCONST TabConstNF (NOLOCK) ON SpNalichF.ID = TabConstNF.OBJID AND TabConstNF.ID = 9792 AND 
                                                    TabConstNF.DATE = (SELECT MAX(TabConstl.DATE) FROM _1SCONST AS TabConstl (NOLOCK)
                                                                WHERE TabConstl.OBJID = TabConstNF.OBJID AND TabConstl.ID = TabConstNF.ID and TabConstl.DATE <=@PeriodK)
                                                where (SpNalichF.IsMark = 0) and (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                                                                 //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                                    ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                                    ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                                    ";";
                if (S_PrihRas)
                {
                    queryString = queryString + "select R.nomk, SpEd.SP78, R.NachOst, R.Prihod, R.Rashod, R.KonOst, R.NachKommis, R.KonKommis, R.NachUMastera, R.KonUMastera, R.NachMastersk, R.KonMastersk, R.PrihVnesh, R.RashVnutr, R.RashPartii, R.ZamenaPlus, R.ZamenaMinus, " +
                                                "R.NachGarZaver, R.KonGarZaver " +
                                                PrepareZaprosDop(startDate, IsTA) +
                                            @"inner join sc84 as SpNom (NOLOCK) on (SpNom.id = R.nomk) 
                                        inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842    
                                        where (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                              //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                                ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                                ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                                ";";
                }
                if (S_Filial)
                {
                    queryString = queryString + "select R.nomk, SpEd.SP78, R.ProdanoM, R.VozvratM, " + ((IsMaster > 0) ? "R.PoGarantiiM,R.PlatnoM,R.PodZakaz,R.PodZakazPlatno, " : "") +
                        "R.PredZakaz,R.Spros,R.SprosOborot,R.RezRaspred,R.VNabore,R.VZayavke,R.OPrihod " + (((listSkl != null) && (listSkl.Count > 0)) ? ",R.Brak " : "") + (S_PrihRas ? "" : ",R.Dostavka,R.Komis,R.UMastera,R.VMastersk ") +
                                                 (((PpBrak != null) && (PpBrak.Count > 0)) ? ",R.PpBrak " : "") +
                                                 (((Brak != null) && (Brak.Count > 0)) ? ",R.BrakMX " : "") +
                                                 (((Roz != null) && (Roz.Count > 0)) ? ",R.Roz " : "") +
                                                 (((Otstoy != null) && (Otstoy.Count > 0)) ? ",R.Otstoy " : "") +
                                                 (((Akt != null) && (Akt.Count > 0)) ? ",R.Akt " : "") +
                                                 (((SobNugdy != null) && (SobNugdy.Count > 0)) ? ",R.SobNugdy " : "") +
                                                 (((Podarky != null) && (Podarky.Count > 0)) ? ",R.Podarky " : "") +
                                                 (((Prokat != null) && (Prokat.Count > 0)) ? ",R.Prokat " : "") +
                                                 (((Ilin != null) && (Ilin.Count > 0)) ? ",R.Ilin " : "") +
                                                PrepareZaprosFilial(startDate, IsTA, (IsMaster > 0), S_PrihRas, listSkl, PpBrak, Brak, Roz, Otstoy, Akt, SobNugdy, Podarky, Prokat, Ilin) +
                                            @"inner join sc84 as SpNom (NOLOCK) on (SpNom.id = R.nomk) 
                                            inner join SC75 as SpEd (NOLOCK) on (SpEd.id = SpNom.SP94)
                                            inner join SC8840 as proiz (NOLOCK) on proiz.id = SpNom.SP8842    
                                            where (SpNom.IsMark = 0) and (SpNom.IsFolder = 2) " + //and " +
                                                                                                  //SprInclude("sc84", "SpNom", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                                ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "SpNom", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                                ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                                ";";

                }
                queryString = queryString + @"select nomenk.id, SpEd.SP78,
                                        nomenk.code, nomenk.descr, IsNull(nomenk.SP8848,''),
                                        IsNull(proiz.descr,''), IsNull(nomenk.SP85,''),
                                        IsNull(Gr1.id, ''), IsNull(Gr2.id, ''), IsNull(Gr3.id, ''),
                                        IsNull(Gr1.sp95,''),IsNull(Gr2.sp95,''),IsNull(Gr3.sp95,''),
                                        IsNull(TabConstWeb.Value, 0), nomenk.SP208, nomenk.SP10397,
                                        IsNull(R_Part.ostatok, 0), IsNull(R_Part.summaUpr, 0)";
                if (S_CostZak)
                    queryString = queryString + @", IsNull(TabConstZak.Value, 0), IsNull(TabConstOpt.Value, 0), IsNull(TabConst4.Value, 0), IsNull(TabConstRoz.Value, 0), IsNull(TabConst6.Value, 0) ";
                queryString = queryString + @", nomenk.SP10366 as KolUpak, nomenk.SP10091 as KolPeremes, nomenk.SP10406 as KolPeremesOpt, IsNull(nomenk.SP9304 ,'') as CertifNum, nomenk.SP9305 as CertifDate,
                                        nomenk.SP5066 as TolkoRozn, nomenk.SP13277 as YandexMarket, (select count(*) from SC552 as SprAnalog (NOLOCK) where nomenk.id = SprAnalog.parentext and SprAnalog.IsMark = 0) as ContainAnalog
                                        from sc84 as nomenk (NOLOCK)
                                        inner join SC75 as SpEd (NOLOCK) on (SpEd.id = nomenk.SP94)
                                        inner join SC8840 as proiz (NOLOCK) on proiz.id = nomenk.SP8842
                                        LEFT OUTER JOIN sc84 As Gr3 (NOLOCK) ON nomenk.PARENTID = Gr3.ID
                                        LEFT OUTER JOIN sc84 As Gr2 (NOLOCK) ON Gr3.PARENTID = Gr2.ID
                                        LEFT OUTER JOIN sc84 As Gr1 (NOLOCK) ON Gr2.PARENTID = Gr1.ID 
                                        LEFT JOIN _1SCONST TabConstWeb (NOLOCK) ON nomenk.ID = TabConstWeb.OBJID AND TabConstWeb.ID = 12305 AND 
                                            TabConstWeb.DATE = (SELECT MAX(TabConstWebl.DATE) FROM _1SCONST AS TabConstWebl (NOLOCK)
                                                                WHERE TabConstWebl.OBJID = TabConstWeb.OBJID AND TabConstWebl.ID = TabConstWeb.ID)
                                        left join (
                                                    select SP331, Sum(SP342) as ostatok, Sum(SP421) as summaUpr from RG328 (NOLOCK)
                                                        where period = @Period and SP340 in ('   2TB   ','   2TC   ') and SP4061 in " + OwnFirms.InCondition + @"
                                                    group by SP331) as R_Part on R_Part.SP331 = nomenk.id ";
                if (S_CostZak)
                    queryString = queryString + @"
                                        left join SC319 as SpCostZak (NOLOCK) on ((nomenk.ID = SpCostZak.PARENTEXT) AND SpCostZak.SP327 = @CostZakType and (SpCostZak.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstZak (NOLOCK) ON SpCostZak.ID = TabConstZak.OBJID AND TabConstZak.ID = 324 AND 
                                            TabConstZak.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstZak.OBJID and id = TabConstZak.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCost (NOLOCK) on ((nomenk.ID = SpCost.PARENTEXT) AND SpCost.SP327 = @CostType and (SpCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstOpt (NOLOCK) ON SpCost.ID = TabConstOpt.OBJID AND TabConstOpt.ID = 324 AND 
                                            TabConstOpt.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstOpt.OBJID and id = TabConstOpt.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCostRoz (NOLOCK) on ((nomenk.ID = SpCostRoz.PARENTEXT) AND SpCostRoz.SP327 = @CostRoznType and (SpCostRoz.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConstRoz (NOLOCK) ON SpCostRoz.ID = TabConstRoz.OBJID AND TabConstRoz.ID = 324 AND 
                                            TabConstRoz.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConstRoz.OBJID and id = TabConstRoz.id
                                                                  order by date desc, time desc)
                                        left join SC319 as SpCostSP (NOLOCK) on ((nomenk.ID = SpCostSP.PARENTEXT) and (SpCostSP.SP327 = @CostSPType) and (SpCostSP.IsMark = 0))
                                        left join SC8904 as SpSklCost (NOLOCK) on ((SpCostSP.ID = SpSklCost.PARENTEXT) and (SpSklCost.SP8901 = @Ekran) and (SpSklCost.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst4 (NOLOCK) ON SpSklCost.ID = TabConst4.OBJID AND TabConst4.ID = 8902 AND 
                                            TabConst4.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst4.OBJID and id = TabConst4.id
                                                                order by date desc, time desc)
                                        left join SC8904 as SpSklCostR (NOLOCK) on ((SpCostSP.ID = SpSklCostR.PARENTEXT) and (SpSklCostR.SP8901 = @SklR) and (SpSklCostR.IsMark = 0))
                                        LEFT OUTER JOIN _1SCONST TabConst6 (NOLOCK) ON SpSklCostR.ID = TabConst6.OBJID AND TabConst6.ID = 8902 AND 
                                           TabConst6.row_id = (select top 1 row_id from _1sconst (NOLOCK) where objid = TabConst6.OBJID and id = TabConst6.id
                                                               order by date desc, time desc) ";
                queryString = queryString + @"
                                        where nomenk.ismark = 0 and nomenk.isfolder = 2 " + //and " +
                                                                                            //SprInclude("sc84", "nomenk", "'" + Nomenklatura.ZapChasti + "'", ((IsMaster > 0) ? true : false)) +
                                            ((listNomk.Count > 0) ? " and " + SprInclude("sc84", "nomenk", string.Join(",", listNomk.Select(x => "'" + x + "'").ToArray()), ((NomenkInclude > 0) ? true : false)) : "") +
                                            ((listProiz.Count > 0) ? " and " + SprInclude("sc8840", "proiz", string.Join(",", listProiz.Select(x => "'" + x + "'").ToArray()), ((ProizInclude > 0) ? true : false), true) : "") +
                                           "order by proiz.descr, nomenk.descr";
                SqlCommand command = new SqlCommand(queryString, connection);
                command.CommandTimeout = 2000;
                command.Parameters.Add("@Period", SqlDbType.DateTime);
                DateTime KonDate = new DateTime(endDate.Year, endDate.Month, 1);
                command.Parameters["@Period"].Value = KonDate;
                if (!IsTA)
                {
                    DateTime PeriodR = new DateTime(startDate.AddMonths(-1).Year, startDate.AddMonths(-1).Month, 1);
                    command.Parameters.Add("@PeriodR", SqlDbType.DateTime);
                    command.Parameters["@PeriodR"].Value = PeriodR;
                    if (startDate != new DateTime(startDate.Year, startDate.Month, 1))
                    {
                        DateTime PeriodN1 = new DateTime(startDate.Year, startDate.Month, 1);
                        command.Parameters.Add("@PeriodN1", SqlDbType.DateTime);
                        command.Parameters["@PeriodN1"].Value = PeriodN1;
                        command.Parameters.Add("@PeriodK1", SqlDbType.DateTime);
                        command.Parameters["@PeriodK1"].Value = startDate.AddDays(1); ;
                    }
                }
                command.Parameters.Add("@PeriodN", SqlDbType.DateTime);
                command.Parameters.Add("@PeriodK", SqlDbType.DateTime);
                command.Parameters["@PeriodN"].Value = startDate;
                command.Parameters["@PeriodK"].Value = endDate.AddDays(1);
                if (S_CostZak)
                {
                    command.Parameters.AddWithValue("@Ekran", Sklad.Ekran); //оптовый склад Экран
                    command.Parameters.AddWithValue("@CostZakType", CostType.Zak);  //тип цен Закупочная
                    command.Parameters.AddWithValue("@CostType", CostType.Opt);  //тип цен Оптовая
                    command.Parameters.AddWithValue("@CostRoznType", CostType.Rozn);  //тип цен Розничная
                    command.Parameters.AddWithValue("@CostSPType", CostType.SP); //тип цен Сп
                    command.Parameters.AddWithValue("@SklR", Sklad.Gastello); //розничный склад Гастелло-Инструмент
                }
                if (S_Filial)
                {
                    command.Parameters.Add("@PeriodNM", SqlDbType.DateTime);
                    command.Parameters.Add("@PeriodNW", SqlDbType.DateTime);
                    command.Parameters["@PeriodNM"].Value = endDate.AddDays(-31);
                    command.Parameters["@PeriodNW"].Value = endDate.AddDays(-7);
                }

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                List<NomenklSkladData> NomSkl = new List<NomenklSkladData>();
                List<NomenklData> NomList = new List<NomenklData>();
                List<Marketplace> Marketplaces = new List<Marketplace>();
                List<NomenkMarketplace> NomenkMarketplaces = new List<NomenkMarketplace>();
                List<ZakazReglament> rez = new List<ZakazReglament>();
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            (SklProperties.FirstOrDefault(n => n.Code == reader.GetString(1).Trim()) ?? new SkladProperties { Code = reader.GetString(1).Trim() }).ID = reader.GetString(0).Trim();
                            (SklProperties.FirstOrDefault(n => n.Code == reader.GetString(1).Trim()) ?? new SkladProperties { Code = reader.GetString(1).Trim() }).Name = reader.GetString(2).Trim();
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                Marketplaces.Add(new Marketplace 
                                {
                                    Id = reader["Id"].ToString(),
                                    Name = reader["Name"].ToString().Trim(),
                                    ShortName = reader["ShortName"].ToString().Trim(),
                                    Parent = reader["Parent"].ToString(),
                                    MarketType = reader["MarketType"].ToString().Trim(),
                                    Model = reader["Model"].ToString().Trim(),
                                });
                            }
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            NomenkMarketplaces.Add(new NomenkMarketplace
                            {
                                NomenkId = reader["nomId"].ToString().Trim(),
                                MarketId = reader["marketId"].ToString(),
                                Comission = Convert.ToDouble(reader["comission"], CultureInfo.InvariantCulture)
                            });
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                            NomenklData nd = (NomList.FirstOrDefault(n => n.NomID == reader.GetString(0).Trim()) ?? new NomenklData { NomID = reader.GetString(0).Trim() });
                            nd.ReglamentTZ = Math.Round(Convert.ToDouble(reader["regTZ"], CultureInfo.InvariantCulture) / koff, 3);
                            nd.OstatokTZ = Math.Round(Convert.ToDouble(reader["Ostatok"], CultureInfo.InvariantCulture) / koff, 3);
                            if (!NomList.Any(x => x.NomID == reader.GetString(0).Trim()))
                                NomList.Add(nd);
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                            NomenklSkladData nsd = new NomenklSkladData();
                            nsd.NomID = reader.GetString(0).Trim();
                            nsd.SklID = reader.GetString(2).Trim();
                            nsd.OstatokN = Math.Round(Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture) / koff, 3);
                            nsd.OstatokK = Math.Round(Convert.ToDouble(reader.GetValue(6), CultureInfo.InvariantCulture) / koff, 3);
                            nsd.OstNabN = Math.Round(Convert.ToDouble(reader["NabN"], CultureInfo.InvariantCulture) / koff, 3);
                            nsd.OstNabK = Math.Round(Convert.ToDouble(reader["NabK"], CultureInfo.InvariantCulture) / koff, 3);
                            nsd.OstRezN = Math.Round(Convert.ToDouble(reader["RezN"], CultureInfo.InvariantCulture) / koff, 3);
                            nsd.OstRezK = Math.Round(Convert.ToDouble(reader["RezK"], CultureInfo.InvariantCulture) / koff, 3);
                            NomSkl.Add(nsd);
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                            NomenklSkladData nsd = (NomSkl.FirstOrDefault(n => n.NomID == reader.GetString(0).Trim() && n.SklID == reader.GetString(2).Trim()) ?? new NomenklSkladData { NomID = reader.GetString(0).Trim(), SklID = reader.GetString(2).Trim() });
                            nsd.MinDocDT = Convert.ToDateTime(reader["MinDocDT"]);
                            nsd.MaxDocDT = Convert.ToDateTime(reader["MaxDocDT"]);
                            nsd.Prodano = Math.Round(Convert.ToDouble(reader["prodano"], CultureInfo.InvariantCulture) / koff, 3) +
                                          Math.Round(Convert.ToDouble(reader["PoGarantii"], CultureInfo.InvariantCulture) / koff, 3) +
                                          Math.Round(Convert.ToDouble(reader["Platno"], CultureInfo.InvariantCulture) / koff, 3);
                            nsd.ProdanoSt = Math.Round(Convert.ToDouble(reader["prStoimost"], CultureInfo.InvariantCulture), 3) +
                                          Math.Round(Convert.ToDouble(reader["PoGarantiiSt"], CultureInfo.InvariantCulture), 3) +
                                          Math.Round(Convert.ToDouble(reader["PlatnoSt"], CultureInfo.InvariantCulture), 3);
                            nsd.SebeSt = Math.Round(Convert.ToDouble(reader["sebest"], CultureInfo.InvariantCulture), 3) +
                                          Math.Round(Convert.ToDouble(reader["PoGarantiiSebeSt"], CultureInfo.InvariantCulture), 3) +
                                          Math.Round(Convert.ToDouble(reader["PlatnoSebeSt"], CultureInfo.InvariantCulture), 3);
                            nsd.Remont = Math.Round(Convert.ToDouble(reader["PoGarantii"], CultureInfo.InvariantCulture) / koff, 3) +
                                        Math.Round(Convert.ToDouble(reader["Platno"], CultureInfo.InvariantCulture) / koff, 3);

                            if (!NomSkl.Any(x => x.NomID == reader.GetString(0).Trim() && x.SklID == reader.GetString(2).Trim()))
                                NomSkl.Add(nsd);
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                            NomenklData nd = (NomList.FirstOrDefault(n => n.NomID == reader.GetString(0).Trim()) ?? new NomenklData { NomID = reader.GetString(0).Trim() });
                            nd.OstatokAnalog = nd.OstatokAnalog + Math.Round(Convert.ToDouble(reader["Ostatok"], CultureInfo.InvariantCulture) / koff, 3);
                            if (!NomList.Any(x => x.NomID == reader.GetString(0).Trim()))
                                NomList.Add(nd);
                        }
                    }
                    if (reader.NextResult())
                    {
                        while (reader.Read())
                        {
                            NomenklSkladData nsd = (NomSkl.FirstOrDefault(n => n.NomID == reader["nomk"].ToString().Trim() && n.SklID == reader["sklad"].ToString().Trim()) ?? new NomenklSkladData { NomID = reader["nomk"].ToString().Trim(), SklID = reader["sklad"].ToString().Trim() });
                            nsd.NalFil = Math.Round(Convert.ToDouble(reader["nalFil"], CultureInfo.InvariantCulture) * Convert.ToDouble(reader["koff"], CultureInfo.InvariantCulture), 2);
                            if (!NomSkl.Any(x => x.NomID == reader["nomk"].ToString().Trim() && x.SklID == reader["sklad"].ToString().Trim()))
                                NomSkl.Add(nsd);
                        }
                    }
                    if (S_PrihRas)
                    {
                        if (reader.NextResult()) //dop
                        {
                            while (reader.Read())
                            {
                                double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                                NomenklData nd = (NomList.FirstOrDefault(n => n.NomID == reader.GetString(0).Trim()) ?? new NomenklData { NomID = reader.GetString(0).Trim() });

                                nd.OstatokDostavkN = Math.Round(Convert.ToDouble(reader.GetValue(2), CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokDostavkK = Math.Round(Convert.ToDouble(reader.GetValue(5), CultureInfo.InvariantCulture) / koff, 3);

                                nd.OstatokKommisN = Math.Round(Convert.ToDouble(reader.GetValue(6), CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokKommisK = Math.Round(Convert.ToDouble(reader.GetValue(7), CultureInfo.InvariantCulture) / koff, 3);

                                nd.OstatokUMasteraN = Math.Round(Convert.ToDouble(reader.GetValue(8), CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokUMasteraK = Math.Round(Convert.ToDouble(reader.GetValue(9), CultureInfo.InvariantCulture) / koff, 3);

                                nd.OstatokMasterskN = Math.Round(Convert.ToDouble(reader.GetValue(10), CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokMasterskK = Math.Round(Convert.ToDouble(reader.GetValue(11), CultureInfo.InvariantCulture) / koff, 3);

                                nd.Prihod = Math.Round(Convert.ToDouble(reader.GetValue(12), CultureInfo.InvariantCulture) / koff, 3) -
                                            Math.Round(Convert.ToDouble(reader.GetValue(13), CultureInfo.InvariantCulture) / koff, 3);
                                nd.Rashod = Math.Round(Convert.ToDouble(reader["RashPartii"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.ZamenaPlus = Math.Round(Convert.ToDouble(reader["ZamenaPlus"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.ZamenaMinus = Math.Round(Convert.ToDouble(reader["ZamenaMinus"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokGarZaverN = Math.Round(Convert.ToDouble(reader["NachGarZaver"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.OstatokGarZaverK = Math.Round(Convert.ToDouble(reader["KonGarZaver"], CultureInfo.InvariantCulture) / koff, 3);
                                if (!NomList.Any(x => x.NomID == reader.GetString(0).Trim()))
                                    NomList.Add(nd);
                            }
                        }
                    }
                    if (S_Filial)
                    {
                        if (reader.NextResult())
                            while (reader.Read())
                            {
                                double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);
                                NomenklData nd = (NomList.FirstOrDefault(n => n.NomID == reader.GetString(0).Trim()) ?? new NomenklData { NomID = reader.GetString(0).Trim() });
                                nd.ProdanoAllM = Math.Round(Convert.ToDouble(reader["ProdanoM"], CultureInfo.InvariantCulture) / koff, 3) -
                                                 Math.Round(Convert.ToDouble(reader["VozvratM"], CultureInfo.InvariantCulture) / koff, 3);
                                if (IsMaster > 0)
                                {
                                    nd.ProdanoAllM = nd.ProdanoAllM + Math.Round(Convert.ToDouble(reader["PoGarantiiM"], CultureInfo.InvariantCulture) / koff, 3) +
                                                                      Math.Round(Convert.ToDouble(reader["PlatnoM"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.RemontM = Math.Round(Convert.ToDouble(reader["PoGarantiiM"], CultureInfo.InvariantCulture) / koff, 3) +
                                                 Math.Round(Convert.ToDouble(reader["PlatnoM"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.PodZakaz = Math.Round(Convert.ToDouble(reader["PodZakaz"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.PodZakazPlatno = Math.Round(Convert.ToDouble(reader["PodZakazPlatno"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.Spros = Math.Round(Convert.ToDouble(reader["Spros"], CultureInfo.InvariantCulture) / koff, 3);
                                }
                                else
                                    nd.Spros = Math.Round(Convert.ToDouble(reader["SprosOborot"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.PredZakaz = Math.Round(Convert.ToDouble(reader["PredZakaz"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.RezRaspred = Math.Round(Convert.ToDouble(reader["RezRaspred"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.VNabore = Math.Round(Convert.ToDouble(reader["VNabore"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.VZayavke = Math.Round(Convert.ToDouble(reader["VZayavke"], CultureInfo.InvariantCulture) / koff, 3);
                                nd.OPrihod = Math.Round(Convert.ToDouble(reader["OPrihod"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((listSkl != null) && (listSkl.Count > 0))
                                    nd.OstatokBrak = Math.Round(Convert.ToDouble(reader["Brak"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((PpBrak != null) && (PpBrak.Count > 0))
                                    nd.OstatokPP = Math.Round(Convert.ToDouble(reader["PpBrak"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Brak != null) && (Brak.Count > 0))
                                    nd.OstatokBrakMX = Math.Round(Convert.ToDouble(reader["BrakMX"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Roz != null) && (Roz.Count > 0))
                                    nd.OstatokRoz = Math.Round(Convert.ToDouble(reader["Roz"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Otstoy != null) && (Otstoy.Count > 0))
                                    nd.OstatokOtstoy = Math.Round(Convert.ToDouble(reader["Otstoy"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Akt != null) && (Akt.Count > 0))
                                    nd.OstatokAkt = Math.Round(Convert.ToDouble(reader["Akt"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((SobNugdy != null) && (SobNugdy.Count > 0))
                                    nd.OstatokSobNugdy = Math.Round(Convert.ToDouble(reader["SobNugdy"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Podarky != null) && (Podarky.Count > 0))
                                    nd.OstatokPodarky = Math.Round(Convert.ToDouble(reader["Podarky"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Prokat != null) && (Prokat.Count > 0))
                                    nd.OstatokProkat = Math.Round(Convert.ToDouble(reader["Prokat"], CultureInfo.InvariantCulture) / koff, 3);
                                if ((Ilin != null) && (Ilin.Count > 0))
                                    nd.OstatokIlin = Math.Round(Convert.ToDouble(reader["Ilin"], CultureInfo.InvariantCulture) / koff, 3);
                                if (!S_PrihRas)
                                {
                                    nd.OstatokDostavkK = Math.Round(Convert.ToDouble(reader["Dostavka"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.OstatokKommisK = Math.Round(Convert.ToDouble(reader["Komis"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.OstatokUMasteraK = Math.Round(Convert.ToDouble(reader["UMastera"], CultureInfo.InvariantCulture) / koff, 3);
                                    nd.OstatokMasterskK = Math.Round(Convert.ToDouble(reader["VMastersk"], CultureInfo.InvariantCulture) / koff, 3);
                                }

                                if (!NomList.Any(x => x.NomID == reader.GetString(0).Trim()))
                                    NomList.Add(nd);
                            }
                    }
                    if (reader.NextResult())
                    {
                        List<string> UsedNomID = new List<string>();
                        
                        while (reader.Read())
                        {
                            string NomID = reader.GetString(0).Trim();
                            if (!UsedNomID.Contains(NomID))
                            {
                                UsedNomID.Add(NomID);
                                double koff = Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture);

                                ZakazReglament zr = new ZakazReglament();
                                zr.NomID = NomID;

                                string Adr1 = "00";
                                string Adr2 = "00";
                                string Adr3 = "00";

                                string Gr1 = reader.GetString(7).Trim();
                                string Gr2 = reader.GetString(8).Trim();
                                string Gr3 = reader.GetString(9).Trim();
                                string Gr1Kom = reader.GetString(10).Trim();
                                string Gr2Kom = reader.GetString(11).Trim();
                                string Gr3Kom = reader.GetString(12).Trim();

                                if (!String.IsNullOrEmpty(Gr1))
                                {
                                    if (Gr3Kom.Length == 2)
                                        Adr3 = Gr3Kom;
                                    if (Gr2Kom.Length == 2)
                                        Adr2 = Gr2Kom;
                                    if (Gr1Kom.Length == 2)
                                        Adr1 = Gr1Kom;
                                }
                                else if (!String.IsNullOrEmpty(Gr2))
                                {
                                    if (Gr2Kom.Length == 2)
                                        Adr1 = Gr2Kom;
                                    if (Gr3Kom.Length == 2)
                                        Adr2 = Gr3Kom;
                                }
                                else if (!String.IsNullOrEmpty(Gr3))
                                {
                                    if (Gr3Kom.Length == 2)
                                        Adr1 = Gr1Kom;
                                }
                                zr.Address = Adr1 + "." + Adr2 + "." + Adr3;
                                if (flHarak == 1)
                                    zr.Articul = Regex.Replace(reader.GetString(4).Trim(), @"\t|\n|\r", ""); 
                                else
                                    zr.Articul = Regex.Replace(reader.GetString(6).Trim(), @"\t|\n|\r", ""); 
                                zr.Brend = Regex.Replace(reader.GetString(5).Trim(), @"\t|\n|\r", "");
                                zr.Nomenklatura = Regex.Replace(reader.GetString(3).Trim(), @"\t|\n|\r", "");

                                List<Marketplace> l_marketplaces = new List<Marketplace>();
                                foreach (var market in Marketplaces)
                                {
                                    l_marketplaces.Add(new Marketplace
                                    {
                                        Name = market.Name,
                                        ShortName = market.ShortName,
                                        Parent = market.Parent,
                                        MarketType = market.MarketType,
                                        Model = market.Model,
                                        Used = NomenkMarketplaces.Exists(x => (x.NomenkId == NomID) && (x.MarketId == market.Id)) ? "*" : "",
                                        Comission = NomenkMarketplaces.Where(x => (x.NomenkId == NomID) && (x.MarketId == market.Id)).Sum(x => x.Comission)
                                    });
                                }
                                zr.Mp = l_marketplaces.ToArray();

                                zr.Site = (reader.GetString(13).Trim() == "1"); //флаг Web

                                zr.KolUpak = Math.Round(Convert.ToDouble(reader["KolUpak"], CultureInfo.InvariantCulture), 3);
                                zr.KolPeremes = Math.Round(Convert.ToDouble(reader["KolPeremes"], CultureInfo.InvariantCulture), 3);
                                zr.KolPeremesOpt = Math.Round(Convert.ToDouble(reader["KolPeremesOpt"], CultureInfo.InvariantCulture), 3);
                                zr.Certif = reader["CertifNum"].ToString().Trim();
                                zr.CertifDate = Convert.ToDateTime(reader["CertifDate"], CultureInfo.InvariantCulture);
                                string certif = "нет";
                                if (zr.Certif != "")
                                {
                                    if ((zr.CertifDate > DateTime.MinValue) && (zr.CertifDate < DateTime.Now.AddMonths(1)))
                                        certif = zr.CertifDate.ToShortDateString();
                                    else
                                        certif = "";
                                }
                                zr.Certif = certif;
                                zr.TolkoRozn = Convert.ToBoolean(reader["TolkoRozn"]) ? "*" : "";
                                zr.YandexMarket = Convert.ToBoolean(reader["YandexMarket"]) ? "*" : "";
                                zr.ContainAnalog = Convert.ToBoolean(reader["ContainAnalog"]) ? "*" : "";
                                zr.NeNadoZakup = Convert.ToBoolean(reader["SP10397"]);
                                zr.ReglamentCenter = (NomSkl.Where(x => x.SklID == Sklad.Center.Trim()).GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, NalFil = q.Sum(s => s.NalFil) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).NalFil;

                                if (Convert.ToDouble(reader.GetValue(14), CultureInfo.InvariantCulture) == 1)
                                    zr.NeZakup = "**";
                                else if (Convert.ToDouble(reader.GetValue(15), CultureInfo.InvariantCulture) == 1)
                                    zr.NeZakup = "*";
                                else
                                    zr.NeZakup = "";
                                zr.Spisstoim = (NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, SebeSt = q.Sum(s => s.SebeSt) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).SebeSt;
                                zr.ProdStoim = (NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, ProdanoSt = q.Sum(s => s.ProdanoSt) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).ProdanoSt;
                                if (Convert.ToDouble(reader.GetValue(16), CultureInfo.InvariantCulture) > 0)
                                    zr.Sebestoim = Math.Round(Convert.ToDouble(reader.GetValue(17), CultureInfo.InvariantCulture) / (Convert.ToDouble(reader.GetValue(16), CultureInfo.InvariantCulture) / koff), 2);
                                int lastColumn = 17;
                                if (S_CostZak)
                                {
                                    zr.CostZak = Convert.ToDouble(reader.GetValue(lastColumn + 1), CultureInfo.InvariantCulture);
                                    zr.CostOpt = Convert.ToDouble(reader.GetValue(lastColumn + 2), CultureInfo.InvariantCulture);
                                    zr.CostOptSP = Convert.ToDouble(reader.GetValue(lastColumn + 3), CultureInfo.InvariantCulture);
                                    zr.CostRozn = Convert.ToDouble(reader.GetValue(lastColumn + 4), CultureInfo.InvariantCulture);
                                    zr.CostRoznSP = Convert.ToDouble(reader.GetValue(lastColumn + 5), CultureInfo.InvariantCulture);
                                    lastColumn = lastColumn + 5;
                                }
                                if (S_PrihRas)
                                {
                                    zr.OstatokStartAll = (NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, OstatokN = q.Sum(s => s.OstatokN) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).OstatokN;
                                    NomenklData currentND = (NomList.FirstOrDefault(n => n.NomID == NomID) ?? new NomenklData { });
                                    zr.OstatokStartAll = zr.OstatokStartAll + currentND.OstatokDostavkN + currentND.OstatokKommisN +
                                        currentND.OstatokUMasteraN + currentND.OstatokMasterskN + currentND.OstatokGarZaverN;
                                    zr.PrihodAll = currentND.Prihod;
                                    zr.RashodAll = currentND.Rashod;
                                    zr.ZamenaPlus = currentND.ZamenaPlus;
                                    zr.ZamenaMinus = currentND.ZamenaMinus;
                                    if (IsMaster > 0)
                                    {
                                        zr.Remont = (NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, Remont = q.Sum(s => s.Remont) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).Remont;
                                    }
                                }
                                if (S_Filial)
                                {
                                    NomenklData currentND = (NomList.FirstOrDefault(n => n.NomID == NomID) ?? new NomenklData { });
                                    NomenklSkladData currentNSD = NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, Prodano = q.Sum(s => s.Prodano), ProdanoSt = q.Sum(s => s.ProdanoSt) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { };
                                    zr.ProdanoAll = currentNSD.Prodano;
                                    double CostOne = zr.ProdanoAll > 0 ? Math.Round(currentNSD.ProdanoSt / zr.ProdanoAll, 2) : 0;
                                    if ((endDate - startDate).TotalDays + 1 >= 30)
                                        zr.ProdanoAll = Math.Round(zr.ProdanoAll / (((endDate - startDate).TotalDays + 1) / 30), 2);
                                    zr.ProdanoAllSum = Math.Round(zr.ProdanoAll * CostOne, 2);
                                    zr.ProdanoAllM = currentND.ProdanoAllM;
                                    zr.ProdanoAllMSum = Math.Round(zr.ProdanoAllM * CostOne, 2);
                                    zr.PredZakaz = currentND.PredZakaz;
                                    zr.Spros = currentND.Spros;
                                    zr.Dostavka = currentND.OstatokDostavkK;
                                    zr.RezRaspred = currentND.RezRaspred;
                                    zr.VNabore = currentND.VNabore;
                                    zr.VZayavke = currentND.VZayavke;
                                    zr.UMastera = currentND.OstatokUMasteraK;
                                    zr.OstGarZaver = currentND.OstatokGarZaverK;
                                    zr.OPrihod = currentND.OPrihod;
                                    zr.NalFilAll = (NomSkl.Where(x => x.SklID != Sklad.Ekran.Trim()).GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, NalFil = q.Sum(s => s.NalFil) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).NalFil;
                                    zr.OstatokAnalog = currentND.OstatokAnalog;
                                    zr.PodZakaz = currentND.PodZakaz;
                                    zr.PodZakazPlatno = currentND.PodZakazPlatno;
                                    zr.ReglamentTZ = currentND.ReglamentTZ;
                                    zr.OstatokTZ = currentND.OstatokTZ;
                                    zr.RemontM = currentND.RemontM;
                                    if (FullInterface)
                                    {
                                        zr.OstatokEndAll = (NomSkl.GroupBy(o => o.NomID).Select(q => new NomenklSkladData { NomID = q.Key, OstatokK = q.Sum(s => s.OstatokK) }).FirstOrDefault(n => n.NomID == NomID) ?? new NomenklSkladData { }).OstatokK;
                                        zr.OstatokEndAll = zr.OstatokEndAll + currentND.OstatokDostavkK;
                                        zr.OstatokEndSvobodAll = zr.OstatokEndAll - zr.VZayavke - currentND.OstatokBrak - zr.VNabore;
                                        zr.OstatokEndAll = zr.OstatokEndAll + currentND.OstatokKommisK +
                                            currentND.OstatokUMasteraK + currentND.OstatokMasterskK;
                                        zr.OstatokEndAllSum = Math.Round(zr.OstatokEndAll * zr.Sebestoim, 2);
                                        zr.OstatokEndSvobodAllSum = Math.Round(zr.OstatokEndSvobodAll * zr.Sebestoim, 2);
                                        zr.Komiss = currentND.OstatokKommisK;
                                        zr.OstHranenie = currentND.OstatokBrak + currentND.OstatokMasterskK;
                                        zr.OstatokPP = currentND.OstatokPP + currentND.OstatokMasterskK;
                                        zr.OstatokBrakMX = currentND.OstatokBrakMX;
                                        zr.OstatokRoz = currentND.OstatokRoz;
                                        zr.OstatokOtstoy = currentND.OstatokOtstoy;
                                        zr.OstatokAkt = currentND.OstatokAkt;
                                        zr.OstatokSobNugdy = currentND.OstatokSobNugdy;
                                        zr.OstatokPodarky = currentND.OstatokPodarky;
                                        zr.OstatokProkat = currentND.OstatokProkat;
                                        zr.OstatokIlin = currentND.OstatokIlin;
                                    }
                                }
                                if (!zr.NeNadoZakup || zr.PodZakaz > 0)
                                {
                                    zr.zakaz = Math.Max((zr.NalFilAll - zr.OstatokEndSvobodAll - zr.OPrihod + (IsMaster > 0 ? zr.PodZakaz : 0)), 0);
                                    zr.zakaz = Math.Ceiling(zr.zakaz);
                                    if (zr.KolUpak > 0)
                                        zr.zakaz = Math.Ceiling(zr.zakaz / zr.KolUpak);
                                }
                                List<SkladProperties> LsklProp = new List<SkladProperties>();
                                foreach (SkladProperties sp in SklProperties)
                                {
                                    SkladProperties Asp = new SkladProperties();
                                    NomenklSkladData nsd = NomSkl.FirstOrDefault(x => x.NomID == zr.NomID && x.SklID == sp.ID) ?? new NomenklSkladData { };
                                    Asp.ID = sp.ID;
                                    Asp.Code = sp.Code;
                                    Asp.Name = sp.Name;
                                    Asp.Ostatok = nsd.OstatokK - nsd.OstNabK - nsd.OstRezK;
                                    Asp.prodagi = nsd.Prodano;
                                    DateTime stDate = startDate;
                                    DateTime edDate = endDate;
                                    if ((nsd.OstatokN == 0) && (nsd.OstatokK > 0) && (nsd.MinDocDT > DateTime.MinValue))
                                    {
                                        stDate = nsd.MinDocDT;
                                        edDate = endDate;
                                    }
                                    else if ((nsd.OstatokN > 0) && (nsd.OstatokK == 0) && (nsd.MaxDocDT > DateTime.MinValue))
                                    {
                                        stDate = startDate;
                                        edDate = nsd.MaxDocDT;
                                    }
                                    else if ((nsd.OstatokN == 0) && (nsd.OstatokK == 0) && (nsd.MinDocDT > DateTime.MinValue) && (nsd.MaxDocDT > DateTime.MinValue))
                                    {
                                        stDate = nsd.MinDocDT;
                                        edDate = nsd.MaxDocDT;
                                    }
                                    if ((edDate - stDate).TotalDays + 1 >= 30)
                                        Asp.prodagi = Math.Round(Asp.prodagi / (((edDate - stDate).TotalDays + 1) / 30), 2);
                                    if ((sp.ID == Sklad.Ekran.Trim()) && (IsMaster == 0))
                                        Asp.nalFil = zr.ReglamentCenter;
                                    else
                                        Asp.nalFil = nsd.NalFil;
                                    LsklProp.Add(Asp);
                                }
                                zr.sp = LsklProp.ToArray();

                                rez.Add(zr);
                            }
                        }
                    }
                    reply.Result = new Results();
                    reply.ExcelFile = FileXLS.ReportZakazReglamentXLS(Extension, "ЗаказПоРегламенту", startDate, endDate, FullInterface,
                                            S_CostZak, S_PrihRas, S_Filial, S_Upak, S_MinOtgr, S_Sertif, (IsMaster > 0), SklProperties,
                                            rez, reply.Result, startDate, endDate, (flBrend > 0));

                    reply.ReplyState.State = "Success";
                    reply.ReplyState.Description = "Информация успешно извлечена из БД";

                    reply.Report = rez.ToArray();

                    //File.WriteAllBytes(@"f:\tmp\test.xls", reply.ExcelFile);
                }
                catch (Exception e)
                {
                    reply.ReplyState.State = "Error";
                    reply.ReplyState.Description = e.Message;
                }
                finally
                {
                    reader.Close();
                }
            }
            return reply;
        }

        [WebMethod(Description = "Создание платежа Ю-Касса")]
        public U_KassaCreatePayment CreateNewPayment(string shopId, string secretKey, string docNo, decimal totalValue, string manager,
            string phone, string email, int taxSystem, List<RItem> rItems)
        {
            U_KassaCreatePayment Response = new U_KassaCreatePayment();
            Response.ReplyState = new ReplyState();
            Response.ReplyState.State = "Error";
            Response.ReplyState.Description = "Ошибка";
            try
            {
                if (string.IsNullOrEmpty(phone) && string.IsNullOrEmpty(email))
                    Response.ReplyState.Description = "Не заданы телефон или email";
                else if (rItems.Count == 0)
                    Response.ReplyState.Description = "Не заданы товары";
                else
                {
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                    var client = new Yandex.Checkout.V3.Client(shopId, secretKey);
                    var newPayment = new Yandex.Checkout.V3.NewPayment
                    {
                        Amount = new Yandex.Checkout.V3.Amount { Value = totalValue, Currency = "RUB" },
                        Description = docNo,
                        Capture = true,
                        Metadata = new Dictionary<string, string>
                        {
                            { "Менеджер", manager }
                        },
                        Confirmation = new Yandex.Checkout.V3.Confirmation
                        {
                            Type = Yandex.Checkout.V3.ConfirmationType.Redirect,
                            ReturnUrl = "https://stinmarket.ru/"
                        },
                        Receipt = new Yandex.Checkout.V3.Receipt
                        {
                            Items = new List<Yandex.Checkout.V3.ReceiptItem>(),
                            TaxSystemCode = (Yandex.Checkout.V3.TaxSystem)taxSystem
                        }
                    };
                    if (!string.IsNullOrEmpty(phone))
                    {
                        if (phone.Length == 10)
                            phone = "7" + phone;
                        else if (phone.Length > 10)
                            phone = "7" + phone.Substring(phone.Length - 10);
                        newPayment.Receipt.Phone = phone;
                    }
                    if (!string.IsNullOrEmpty(email))
                        newPayment.Receipt.Email = email;
                    foreach (var i in rItems)
                    {
                        var item = new Yandex.Checkout.V3.ReceiptItem
                        {
                            Description = i.Description,
                            Quantity = i.Quantity,
                            Amount = new Yandex.Checkout.V3.Amount { Value = i.AmountValue, Currency = "RUB" },
                            VatCode = (Yandex.Checkout.V3.VatCode)i.NdsNo,
                            PaymentMode = Yandex.Checkout.V3.PaymentMode.FullPrepayment, //.FullPayment,
                            PaymentSubject = i.Type == 0 ? Yandex.Checkout.V3.PaymentSubject.Commodity : Yandex.Checkout.V3.PaymentSubject.Service, //.Payment, //
                        };
                        if (!string.IsNullOrEmpty(i.CountryCode))
                            item.CountryOfOriginCode = i.CountryCode;
                        if (!string.IsNullOrEmpty(i.Gtd))
                            item.CustomsDeclarationNumber = i.Gtd;
                        newPayment.Receipt.Items.Add(item);
                    }
                    var payment = client.CreatePayment(newPayment);

                    Response.ReplyState.State = "Success";
                    Response.ReplyState.Description = "Payment успешно создан";
                    Response.PaymentId = payment.Id;
                    Response.ConfirmationUrl = payment.Confirmation.ConfirmationUrl;
                    Response.IntStatus = (int)payment.Status;
                }
            }
            catch (Yandex.Checkout.V3.YandexCheckoutException ye)
            {
                Response.ReplyState.State = "Error";
                Response.ReplyState.Description = ye.Error.Parameter;
                Response.ReplyState.Description += ", " + ye.Error.Description;
            }
            catch (Exception e)
            {
                Response.ReplyState.State = "Error";
                Response.ReplyState.Description = e.Message;
                var st = new System.Diagnostics.StackTrace(e, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                Response.ReplyState.Description += ", " + frame.GetFileLineNumber();
            }
            return Response;
        }
        [WebMethod(Description = "Зачет платежа Ю-Касса")]
        public ReplyState OffsetPayment(string shopId, string secretKey, string paymentId, string customerName, string customerInn,
            string customerPhone, string customerEmail, int taxSystem, List<RItem> rItems, List<RSettlement> rSettlements)
        {
            ReplyState Response = new ReplyState();
            Response.State = "Error";
            Response.Description = "Ошибка";
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                var client = new Yandex.Checkout.V3.Client(shopId, secretKey);
                var receipt = new Yandex.Checkout.V3.SettlementReceipt();
                //receipt.Id = 
                receipt.Type = Yandex.Checkout.V3.SettlementReceiptType.Payment;
                receipt.PaymentId = paymentId;
                //receipt.RefundId =
                if (!string.IsNullOrEmpty(customerPhone))
                {
                    if (customerPhone.Length == 10)
                        customerPhone = "7" + customerPhone;
                    else if (customerPhone.Length > 10)
                        customerPhone = "7" + customerPhone.Substring(customerPhone.Length - 10);
                }

                receipt.Customer = new Yandex.Checkout.V3.Customer()
                {
                    //FullName = string.IsNullOrEmpty(customerName) ? null : customerName,
                    //Inn = string.IsNullOrEmpty(customerInn) ? null : customerInn,
                    Phone = string.IsNullOrEmpty(customerPhone) ? null : customerPhone,
                    Email = string.IsNullOrEmpty(customerEmail) ? null : customerEmail
                };
                receipt.Items = new List<Yandex.Checkout.V3.ReceiptItem>();
                foreach (var i in rItems)
                {
                    var item = new Yandex.Checkout.V3.ReceiptItem
                    {
                        Description = i.Description,
                        Quantity = i.Quantity,
                        Amount = new Yandex.Checkout.V3.Amount { Value = i.AmountValue, Currency = "RUB" },
                        VatCode = (Yandex.Checkout.V3.VatCode)i.NdsNo,
                        PaymentMode = Yandex.Checkout.V3.PaymentMode.FullPayment, //.FullPrepayment,
                        PaymentSubject = Yandex.Checkout.V3.PaymentSubject.Commodity, //.Payment, //
                    };
                    if (!string.IsNullOrEmpty(i.CountryCode))
                        item.CountryOfOriginCode = i.CountryCode;
                    if (!string.IsNullOrEmpty(i.Gtd))
                        item.CustomsDeclarationNumber = i.Gtd;
                    receipt.Items.Add(item);
                }
                receipt.TaxSystemCode = (Yandex.Checkout.V3.TaxSystem)taxSystem;
                receipt.Send = true;
                receipt.Settlements = new List<Yandex.Checkout.V3.Settlement>();
                foreach (var i in rSettlements)
                {
                    var item = new Yandex.Checkout.V3.Settlement
                    {
                        Amount = new Yandex.Checkout.V3.Amount { Value = i.AmountValue, Currency = "RUB" },
                        Type = (Yandex.Checkout.V3.SettlementType)i.Type
                    };
                    receipt.Settlements.Add(item);
                }
                var result = client.CreateSettlementReceipt(receipt);
                if (!string.IsNullOrEmpty(result.Id))
                {
                    Response.State = "Success";
                    Response.Description = result.Id;
                }
            }
            catch (Yandex.Checkout.V3.YandexCheckoutException ye)
            {
                Response.State = "Error";
                Response.Description = ye.Error.Parameter;
                Response.Description += ", " + ye.Error.Description;
            }
            catch (Exception e)
            {
                Response.State = "Error";
                Response.Description = e.Message;
                var st = new System.Diagnostics.StackTrace(e, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                Response.Description += ", " + frame.GetFileLineNumber();
            }
            return Response;
        }
        [WebMethod(Description = "Возврат платежа Ю-Касса")]
        public U_KassaCreateRefund RefundPayment(string shopId, string secretKey, string paymentId, string docNo, decimal totalValue, 
            string customerPhone, string customerEmail, int taxSystem, List<RItem> rItems)
        {
            U_KassaCreateRefund Response = new U_KassaCreateRefund();
            Response.ReplyState = new ReplyState();
            Response.ReplyState.State = "Error";
            Response.ReplyState.Description = "Ошибка";
            try
            {
                if (string.IsNullOrEmpty(customerPhone) && string.IsNullOrEmpty(customerEmail))
                    Response.ReplyState.Description = "Не заданы телефон или email";
                else if (rItems.Count == 0)
                    Response.ReplyState.Description = "Не заданы товары";
                else
                {
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                    var client = new Yandex.Checkout.V3.Client(shopId, secretKey);
                    var newRefund = new Yandex.Checkout.V3.NewRefund
                    {
                        PaymentId = paymentId,
                        Amount = new Yandex.Checkout.V3.Amount { Value = totalValue, Currency = "RUB" },
                        Description = docNo,
                        Receipt = new Yandex.Checkout.V3.Receipt
                        {
                            Items = new List<Yandex.Checkout.V3.ReceiptItem>(),
                            TaxSystemCode = (Yandex.Checkout.V3.TaxSystem)taxSystem
                        }
                    };
                    if (!string.IsNullOrEmpty(customerPhone))
                    {
                        if (customerPhone.Length == 10)
                            customerPhone = "7" + customerPhone;
                        else if (customerPhone.Length > 10)
                            customerPhone = "7" + customerPhone.Substring(customerPhone.Length - 10);
                        newRefund.Receipt.Phone = customerPhone;
                    }
                    if (!string.IsNullOrEmpty(customerEmail))
                        newRefund.Receipt.Email = customerEmail;
                    foreach (var i in rItems)
                    {
                        var item = new Yandex.Checkout.V3.ReceiptItem
                        {
                            Description = i.Description,
                            Quantity = i.Quantity,
                            Amount = new Yandex.Checkout.V3.Amount { Value = i.AmountValue, Currency = "RUB" },
                            VatCode = (Yandex.Checkout.V3.VatCode)i.NdsNo,
                            PaymentMode = Yandex.Checkout.V3.PaymentMode.FullPayment, 
                            PaymentSubject = i.Type == 0 ? Yandex.Checkout.V3.PaymentSubject.Commodity : Yandex.Checkout.V3.PaymentSubject.Service, //.Payment, //
                        };
                        if (!string.IsNullOrEmpty(i.CountryCode))
                            item.CountryOfOriginCode = i.CountryCode;
                        if (!string.IsNullOrEmpty(i.Gtd))
                            item.CustomsDeclarationNumber = i.Gtd;
                        newRefund.Receipt.Items.Add(item);
                    }
                    var refund = client.CreateRefund(newRefund);

                    Response.ReplyState.State = "Success";
                    Response.ReplyState.Description = "Payment успешно создан";
                    Response.RefundId = refund.Id;
                    Response.IntStatus = (int)refund.Status;
                    if (refund.ReceiptRegistration != null)
                        Response.IntReceiptRegistration = (int)refund.ReceiptRegistration;
                }
            }
            catch (Yandex.Checkout.V3.YandexCheckoutException ye)
            {
                Response.ReplyState.State = "Error";
                Response.ReplyState.Description = ye.Error.Parameter;
                Response.ReplyState.Description += ", " + ye.Error.Description;
            }
            catch (Exception e)
            {
                Response.ReplyState.State = "Error";
                Response.ReplyState.Description = e.Message;
                var st = new System.Diagnostics.StackTrace(e, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                Response.ReplyState.Description += ", " + frame.GetFileLineNumber();
            }
            return Response;
        }
    }
}