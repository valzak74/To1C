using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace To1C
{
    class CostType
    {
        //Типы цен
        public const string Zak = "     1   ";  //Закупочная
        public const string Opt = "     2   ";  //Оптовая
        public const string Rozn = "     4   "; //Розничная
        public const string SP = "     6S  ";   //Сп
        public const string Osob = "     PD  ";   //Особая
    }

    class Sklad
    {
        //Склады
        public const string Ekran = "     DS  "; //Экран
        public const string Gastello = "     7S  "; //Гастелло-Инструмент
        public const string Center = "     1   "; //ЦентральныйСклад
        public const string Otstoy = "     6S  "; //Отстой
        public const string DomSad = "     9S  "; //Дом&Сад (Заводское)
    }

    class PodSklad
    {
        public const string TorgZal = "     IS  "; //Торговый зал
    }

    class Nomenklatura
    {
        public const string Musor = "   9U7S  "; //Мусор код S00012726
        public const string ZapChasti = "    17S  "; //Запасные части код S00000018
        public const string ForMaster = "   O8YA  "; //Для мастерской код A00003191
        public const string PryamProd = "  17M7D  "; //Прямые продажи код D00027231
        public const string TorgOborud = "   RQOD  "; //Торговое оборудование и рекламные материалы код D00007946
        public const string Uchenka = "   FXID  "; //17. Уцененные товары код D00000015
    }

    class Firma
    {
        public const string IP_pavlov = "     3S  "; //ИП Павлов
        public const string Stin_service = "     4S  "; //Стин-Сервис
        public const string StinPlus = "     1   "; //СТИН+
    }

    class HDS
    {
        public const string BezHDS = "    I9   "; //без НДС
        public const string HDS_0 = "   9YV   "; //НДС 0%
        public const string HDS_10 = "    I8   "; //НДС 10%
        public const string HDS_18 = "   6F2   "; //НДС 18%
        public const string HDS_20 = "    I7   "; //НДС 20%
    }

    class OwnFirms
    {
        public const string InCondition = "('" + Firma.IP_pavlov + "','" + Firma.StinPlus + "','" + Firma.Stin_service + "')";
    }
}
