using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Informer1C
{
    public class Заявка
    {
        public string IDDOC { get; set; }
        public string НомерДок { get; set; }
        public string ДатаДок { get; set; }
        public string НомерДокПолный { get; set; }
        public string Автор { get; set; }
        public string КонтрагентID { get; set; }
        public string Контрагент { get; set; }
        public List<ЗаявкаСтрока> СтрокиДокумента { get; set; }
    }

    public class ЗаявкаСтрока
    {
        public string НоменклатураID { get; set; }
        public string Номенклатура { get; set; }
        public string Артикул { get; set; }
        public string ПроизводительID { get; set; }
        public string Производитель { get; set; }
        public decimal? Количество { get; set; }
        public decimal? ЦенаЗаявленная { get; set; }
        public decimal? ЦенаСогласованная { get; set; }
        public int? СрокПоставки { get; set; }
        public Вариант Действие { get; set; }
    }
    
    public enum Вариант 
    {
        НеОбработана,
        Отказ,
        Согласована 
    };

    public class НоваяЗаявка
    {
        public string IDDOC { get; set; }
        public string Контрагент { get; set; }
        public string НомерДокПолный { get; set; }
    }

    public class FirmaDetails
    {
        public string Id;
        public string Name;
        public string UrLitso;
        public string Prefix;
        public string BankSchet;
    }

}
