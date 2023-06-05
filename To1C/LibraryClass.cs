using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace To1C
{
    public class TableData
    {
        public string reg { get; set; }
        public int curTable { get; set; }
        public int totalTable { get; set; }
        public List<DDS> paramsDDS { get; set; }
        public List<GroupCondition> conditionsGroup { get; set; }
        public string Cond1 { get; set; }
        public string Cond2 { get; set; }
        private int _type = 0;
        public int Type { get { return _type; } set { _type = value; } } ////Type=0 -- all; =1 -- only rg; =2 -- only ra
        private bool _isMonth = false;
        public bool IsMonth { get { return _isMonth; } set { _isMonth = value; } }
        private bool _isWeek = false;
        public bool IsWeek { get { return _isWeek; } set { _isWeek = value; } }
    }

    public class GroupCondition
    {
        public string name { get; set; }
        public string conditionType { get; set; }
        public List<GroupCondition> subCondition { get; set; }
        public List<DDS> conditionsDDS { get; set; }
    }

    public class DDS
    {
        public string name { get; set; }
        public string conditionType { get; set; }
        public string param { get; set; }
        public string condition { get; set; }
    }

    public class SkladProperties
    {
        public string ID { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public double Ostatok { get; set; }
        public double prodagi { get; set; }
        public double nalFil { get; set; }
        public List<SkladProperties> PodSklads { get; set; }
    }
    
    public class NomenklData
    {
        public string NomID { get; set; }
        public double OstatokDostavkN { get; set; }
        public double OstatokDostavkK { get; set; }
        public double OstatokKommisN { get; set; }
        public double OstatokKommisK { get; set; }
        public double OstatokUMasteraN { get; set; }
        public double OstatokUMasteraK { get; set; }
        public double OstatokMasterskN { get; set; }
        public double OstatokMasterskK { get; set; }
        public double OstatokGarZaverN { get; set; }
        public double OstatokGarZaverK { get; set; }
        public double Prihod { get; set; }
        public double Rashod { get; set; }
        public double ZamenaPlus { get; set; }
        public double ZamenaMinus { get; set; }
        public double ProdanoAllM { get; set; }
        public double PredZakaz { get; set; }
        public double Spros { get; set; }
        public double RezRaspred { get; set; }
        public double VNabore { get; set; }
        public double VZayavke { get; set; }
        public double OPrihod { get; set; }
        public double OstatokBrak { get; set; }
        public double OstatokAnalog { get; set; }
        public double PodZakaz { get; set; }
        public double PodZakazPlatno { get; set; }
        public double ReglamentTZ { get; set; }
        public double OstatokTZ { get; set; }
        public double OstatokPP { get; set; }
        public double OstatokBrakMX { get; set; }
        public double OstatokRoz { get; set; }
        public double OstatokOtstoy { get; set; }
        public double OstatokAkt { get; set; }
        public double OstatokSobNugdy { get; set; }
        public double OstatokPodarky { get; set; }
        public double OstatokProkat { get; set; }
        public double OstatokIlin { get; set; }
        public double RemontM { get; set; }
    }

    public class NomenklSkladData
    {
        public string NomID { get; set; }
        public string SklID { get; set; }
        public DateTime MinDocDT { get; set; }
        public DateTime MaxDocDT { get; set; }
        public double OstatokN { get; set; }
        public double OstatokK { get; set; }
        public double OstNabN { get; set; }
        public double OstNabK { get; set; }
        public double OstRezN { get; set; }
        public double OstRezK { get; set; }
        public double Prodano { get; set; }
        public double ProdanoSt { get; set; }
        public double SebeSt { get; set; }
        public double NalFil { get; set; }
        public double Remont { get; set; }
    }

    public class ZakazReglament
    {
        public string NomID { get; set; }
        public string Address { get; set; }
        public string Articul { get; set; }
        public string Nomenklatura { get; set; }
        public string Brend { get; set; }
        public bool Site { get; set; }
        public string NeZakup { get; set; }
        public double Sebestoim { get; set; }
        public double Spisstoim { get; set; }
        public double ProdStoim { get; set; }
        public double CostOpt { get; set; }
        public double CostOptSP { get; set; }
        public double CostRozn { get; set; }
        public double CostRoznSP { get; set; }
        public double CostZak { get; set; }
        public double OstatokStartAll { get; set; }
        public double PrihodAll { get; set; }
        public double RashodAll { get; set; }
        public double ZamenaPlus { get; set; }
        public double ZamenaMinus { get; set; }
        public double ProdanoAll { get; set; }
        public double ProdanoAllSum { get; set; }
        public double ProdanoAllM { get; set; }
        public double ProdanoAllMSum { get; set; }
        public double PredZakaz { get; set; }
        public double Spros { get; set; }
        public double Dostavka { get; set; }
        public double RezRaspred { get; set; }
        public double VNabore { get; set; }
        public double VZayavke { get; set; }
        public double UMastera { get; set; }
        public double OstGarZaver { get; set; }
        public double OPrihod { get; set; }
        public double OstatokEndAll { get; set; }
        public double OstatokEndAllSum { get; set; }
        public double OstatokEndSvobodAll { get; set; }
        public double OstatokEndSvobodAllSum { get; set; }
        public double NalFilAll { get; set; }
        public double OstatokAnalog { get; set; }
        public double PodZakaz { get; set; }
        public double PodZakazPlatno { get; set; }
        public double zakaz { get; set; }
        public double KolUpak { get; set; }
        public bool NeNadoZakup { get; set; }
        public double ReglamentCenter { get; set; }
        public double ReglamentTZ { get; set; }
        public double OstatokTZ { get; set; }
        public double Komiss { get; set; }
        public double OstHranenie { get; set; }
        public double OstatokPP { get; set; }
        public double OstatokBrakMX { get; set; }
        public double OstatokRoz { get; set; }
        public double OstatokOtstoy { get; set; }
        public double OstatokAkt { get; set; }
        public double OstatokSobNugdy { get; set; }
        public double OstatokPodarky { get; set; }
        public double OstatokProkat { get; set; }
        public double OstatokIlin { get; set; }
        public double KolPeremes { get; set; }
        public double KolPeremesOpt { get; set; }
        public string Certif { get; set; }
        public DateTime CertifDate { get; set; }
        public string YandexMarket { get; set; }
        public string TolkoRozn { get; set; }
        public string ContainAnalog { get; set; }
        public double Remont { get; set; }
        public double RemontM { get; set; }
        public SkladProperties[] sp { get; set; }
        public Marketplace[] Mp { get; set; }
    }
    public class Marketplace
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ShortName { get; set; }
        public string Parent { get; set; }
        public string MarketType { get; set; }
        public string Model { get; set; }
        public string Used { get; set; }
        public double Comission { get; set; }
        public double VolumeWeight { get; set; }
    }
    public class NomenkMarketplace
    {
        public string NomenkId { get; set; }
        public string MarketId { get; set; }
        public bool Deleted { get; set; }
        public double Comission { get; set; }
        public double VolumeWeight { get; set; }
   }
    public class NomenkProperties
    {
        public string NomenkID { get; set; }
        public string Property { get; set; }
        public string Data { get; set; }
    }

    public class Price
    {
        public int level { get; set; }
        public string ID { get; set; }
        public string ParentID { get; set; }
        public string ParentID1 { get; set; }
        public string ParentID2 { get; set; }
        public string ParentID3 { get; set; }
        public int IsFolder { get; set; }
        public string Articul { get; set; }
        public string Nomenklatura { get; set; }
        public string Brend { get; set; }
        public string Harakter { get; set; }
        public double Ostatok { get; set; }
        public List<double> OstatokSkl { get; set; }
        public double OstatokPrihod { get; set; }
        public double Cost { get; set; }
        public double SuperCost { get; set; }
        public double SuperCostR { get; set; }
        public double CostOsob { get; set; }
        public double CostRozn { get; set; }
        public double CostClient { get; set; }
        public bool NeNadoZakup { get; set; }
    }

    public class PriceList
    {
        public ReplyState ReplyState { get; set; }
        public byte[] ExcelFile { get; set; }
    }

    public class ZakazReglamentReply
    {
        public ReplyState ReplyState { get; set; }
        public byte[] ExcelFile { get; set; }
        public ZakazReglament[] Report { get; set; }
        public Results Result { get; set; }
    }
    public class StandartReply
    {
        public ReplyState ReplyState { get; set; }
        public byte[] ExcelFile { get; set; }
    }
    public class U_KassaCreatePayment
    {
        public ReplyState ReplyState { get; set; }
        public string PaymentId { get; set; }
        public string ConfirmationUrl { get; set; }
        public int IntStatus { get; set; }
        public string Status
        {
            get { return this.IntStatus == 0 ? "Pending" : this.IntStatus == 1 ? "WaitingForCapture" : this.IntStatus == 2 ? "Succeeded" : "Canceled"; }
        }
    }
    public class U_KassaCreateRefund
    {
        public ReplyState ReplyState { get; set; }
        public string RefundId { get; set; }
        public int IntStatus { get; set; }
        public string Status
        {
            get { return this.IntStatus == 0 ? "Pending" : this.IntStatus == 1 ? "WaitingForCapture" : this.IntStatus == 2 ? "Succeeded" : "Canceled"; }
        }
        public int IntReceiptRegistration { get; set; }
        public string ReceiptRegistration
        {
            get { return this.IntReceiptRegistration == 1 ? "Pending" : this.IntReceiptRegistration == 2 ? "Succeeded" : "Canceled"; }
        }
    }
    public class RItem
    {
        public int Type { get; set; } // 0-товар; 1-услуга
        public string Description { get; set; }
        public decimal Quantity { get; set; }
        public decimal AmountValue { get; set; }
        public int NdsNo { get; set; }
        public string CountryCode { get; set; }
        public string Gtd { get; set; }
    }
    public class RSettlement
    {
        public decimal AmountValue { get; set; }
        public int Type { get; set; }

    }
    public class DataEntry
    {
        public string Имя { get; set; }
        public string Тип { get; set; }
        public string Значение { get; set; }
    }

    public class nDataEntry
    {
        public string мнИмя { get; set; }
        public List<List<DataEntry>> мнЗначения { get; set; }
    }

    public class NomenklaturaProdanaReply
    {
        public ReplyState ReplyState { get; set; }
        public List<CustTovar> Tovars { get; set; }
    }
    public class NomenklaturaFirma
    {
        public string Код { get; set; }
        public string Наименование { get; set; }
        public string ПолнНаименование { get; set; }
        public string Артикул { get; set; }
        public string ВидНоменклатуры { get; set; }
        public string ЕдиницаКод { get; set; }
        public string ЕдиницаНаименование { get; set; }
        public string СтавкаНДС { get; set; }
        public string СтранаКод { get; set; }
        public string СтранаНаименование { get; set; }
        public double Остаток { get; set; }
        public double Себестоимость { get; set; }
        public double ОстатокПринятый { get; set; }
        public double СебестоимостьПринятый { get; set; }
    }
    public class RowYandexFBY
    {
        public string Id { get; set; }
        public string Sku { get; set; }
        public string Артикул { get; set; }
        public string Наименование { get; set; }
        public double Квант { get; set; }
        public double Цена { get; set; }
        public double ЦенаСП { get; set; }
        public double ЦенаЗакуп { get; set; }
        public double СвободныйОстаток { get; set; }
        public double ОжидаемыйПриход { get; set; }
        public double Поставлено { get; set; }
        public double Впути { get; set; }
        public double ОстаткиНаСкладеЯндекс1 { get; set; }
        public bool Allowed1 { get; set; }
        public double НовыйЗаказ1 { get; set; }
        public double ОстаткиНаСкладеЯндекс2 { get; set; }
        public bool Allowed2 { get; set; }
        public double НовыйЗаказ2 { get; set; }
        public double ОстаткиНаСкладеЯндекс3 { get; set; }
        public bool Allowed3 { get; set; }
        public double НовыйЗаказ3 { get; set; }
        public double ЦенаНаЯндексе { get; set; }
        public double ЦенаПервая { get; set; }
    }
    public class RowYandexStock
    {
        public string SKU { get; set; }
        public bool Allowed { get; set; }
        public double Quantity { get; set; }
    }
    public class Results
    {
        public double ProdanoAllSum { get; set; }
        public double ProdanoAllMSum { get; set; }
        public double OstatokEndAllSum { get; set; }
        public double OstatokEndSvobodAllSum { get; set; }
        public double SumReglament { get; set; }
        public double SumZakaz { get; set; }
    }

    public class CompleteRepairStatus
    {
        public string KvitNumber { get; set; }
        public int KvitDate { get; set; }
        public string Kontragent { get; set; }
        public string Izdelie { get; set; }
        public string BrendID { get; set; }
        public string Brend { get; set; }
        public string Postav { get; set; }
        public string PostavID { get; set; }
        public string Manager { get; set; }
        public string ZavNumber { get; set; }
        public string RepairType { get; set; }
        public DateTime DatePriema { get; set; }
        public string Status { get; set; }
        public string Avtor { get; set; }
        public string AvtorID { get; set; }
        public string Master { get; set; }
        public string MasterID { get; set; }
        public string DocZaverId { get; set; }
        public DateTime DateZaver { get; set; }
        public decimal PredvarSumma { get; set; }
        public decimal PredvarSummaZ { get; set; }
        public decimal PredvarSummaR { get; set; }
        public decimal Summa { get; set; }
        public decimal SummaZ { get; set; }
        public decimal SummaR { get; set; }
        public DateTime DateVyd { get; set; }
        public DateTime DatePer { get; set; }
        public DateTime DatePri { get; set; }
        public DateTime DateOtchOtp { get; set; }
        public DateTime DateOtchPri { get; set; }
        public DateTime DateGarOtch { get; set; }
        public decimal SummaGarOtch { get; set; }
    }

    public class ReplyState
    {
        public string State { get; set; }
        public string Description { get; set; }
    }

    public class FullRepairStatus
    {
        public ReplyState ReplyState { get; set; }
        public CompleteRepairStatus CompleteRepairStatus { get; set; }
    }

    public class RepairStatusArray
    {
        public ReplyState ReplyState { get; set; }
        public int KolOpenRemont { get; set; }
        public int KolClosedRemont { get; set; }
        public int KolVydRemont { get; set; }
        public int KolNeVydRemont { get; set; }
        public decimal Summa { get; set; }
        public decimal SummaZ { get; set; }
        public decimal SummaR { get; set; }
        public decimal SrSummaOneRemont { get; set; }
        public double SrTimeOpenRemont { get; set; }
        public double SrTimeClosedRemont { get; set; }
        public double SrTimeVydRemont { get; set; }
        public int KolPerDocument { get; set; }
        public int KolPriDocument { get; set; }
        public decimal SummaPerDocument { get; set; }
        public decimal SummaPriDocument { get; set; }
        public int KolOtchOtp { get; set; }
        public decimal SummaOtchOtp { get; set; }
        public decimal Delta { get; set; }
        public CompleteRepairStatus[] RepairStatuses { get; set; }
    }

    class Doc
    {
        public string Order_type;
        public string Order_numerator;
        public Doc Order_parent;
        public string Order_Firma;
        public string Order_id;
        public DateTime Order_date;
        public string Order_Customer;
        public string Order_Sklad;
        public string Order_PodSklad;
        public string Order_Ref;
        public double Order_TotalSum;
        public List<DocRecord> Records;
    }

    class DocRecord
    {
        public string Product_id;
        public string Product_name;
        public double Quantity;
        public double Summa;
        public double SkidkaNovTovar;
    }

    public class FirmaDetails
    {
        public string Id;
        public string Name;
        public string UrLitso;
        public string Prefix;
        public string Postfix;
        public string BankSchet;
    }

    public class Customer
    {
        public string Id;
        public string Code;
        public string Name;
        public string MainDogovorId;
        public string TypeCostId;
        public string SkidkaId;
        public double SkidkaValue;
        public string GroupId;
        private bool export = false;
        public bool Export { get { return export; } set { export = value; } }
        public FirmaDetails Firma;
        //public string UrLitso;
        //public string BankSchet;
        //public string FirmaName;
        public string KolonkaSkidki;
        public List<ProizSkidka> proizSkidka = new List<ProizSkidka>();
        public List<ProizSkidka> proizSkidkaM = new List<ProizSkidka>();
        public List<string> proizOtsrochka0 = new List<string>();
        public List<string> proizOtsrochka0M = new List<string>();
    }

    public class ProizSkidka
    {
        public string ProizId;
        public double Skidka;
    }

    public class CustTovar
    {
        public string CustCode;
        public string Id;
        public string Code;
        public string Name;
    }
    public class Tovar
    {
        public string Id;
        public string Code;
        public string Name;
        public string EdId;
        public double koff;
        public double sebestoim;
        public double PorogSebestoim;
        public string ProizId;
        public Costs Costs = new Costs();
        public double SkidkaVsem;
    }
    public class Costs
    {
        public double Zak;
        public double Rozn;
        public double Opt;
        public double Osob;
        public double SP;
        public double SPR;
    }

}