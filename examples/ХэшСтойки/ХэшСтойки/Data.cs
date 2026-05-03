using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using TFlex.Model.Data.ProductStructure;

namespace TFAPIHash
{
    public static class Data
    {
        public static string fname12 = "Этаж.grb";
        public static string fname3 = "Изделие3.grb";
        public static string fname4 = "Изделие4.grb";

        public static string dname12 = "Деталь_Стойка.grb";
        public static string dname34 = "Деталь_Стойка2.grb";

        public static string copyDirectoryName = "copy";
        public static string ДеталиDirectoryName = "Стойки";
        public static string ПлоскостиDirectoryName = "Плоскости";
        public static string PDFDirectoryName = "PDF";
        public static string DWGDirectoryName = "DWG";

        public static string СоставИзделия = "m2spec"; // здесь требуется увязка при использовании составов, отличных от...

        public static string плоскостьVarName = "$Плоскость";
        public static string ЭтажVarName = "Этаж";
        public static string СтойкаVarName = "Стойка";
        public static string ТипЧертежаVarName = "$ТипЧертежа";
        public static string ВидимостьVarName = "$Видимость";
        public static string ЛистVarName = "Лист";
        public static string ЛистовVarName = "Листов";
        public static string Лист_ЧертёжVarName = "Лист_Чертёж";
        public static string Лист_ЗаказVarName = "Лист_Заказ";
        public static string Лист_ОтгрузкаVarName = "Лист_Отгрузка";

        public static string Стойка_ДетальVarName = "Стойка_Деталь"; // Настройка отображения стойки для чертежей стоек
        public static string АртикулСтойкиVarName = "$АртикулДетали"; // Получить артикул стойки из фрагмента


        public static string плоскость1 = "(5-1)-(5-6)";
        public static string плоскость2 = "(5-6)-(5-1)";
        public static string плоскость3 = "(5-А)-(5-Л)";
        public static string плоскость4 = "(5-Л)-(5-А)";

        public static List<string> Видимость = new List<string>() { "Заполнения", "Каркас", "Монтаж" };
        

        public static List<string> СтолбцыРаскрой = new List<string>() { "Плоскость", "Этаж", "Артикул", "Наименование", "Длина", "Количество всего" };
        public struct view { public string СоставИзделия; public string Представление; public List<string> columns; public bool OneGroupOneSheet; }

        public static view Раскрой = new view
        {
            СоставИзделия = "m2spec",
            Представление = "Профили (раскрой)",
            columns = new List<string>() { "Артикул", "Наименование", "Длина", "Место установки", "Количество всего" },
            OneGroupOneSheet = false
        };

        public static view ПокупныеИзделия = new view
        {
            СоставИзделия = "m2spec",
            Представление = "Покупные изделия (заказ)",
            columns = new List<string>() {"Артикул", "Наименование", "Место установки", "Количество всего" },
            OneGroupOneSheet = false
        };

        public static view Материалы = new view
        {
            СоставИзделия = "m2spec",
            Представление = "Материалы (заказ)",
            columns = new List<string>() {"Артикул", "Материал", "Место установки", "Значение всего", "Единица измерения" },
            OneGroupOneSheet = false
        };

        public static view Заполнения = new view
        {
            СоставИзделия = "m2spec",
            Представление = "Заполнения (заказ)",
            columns = new List<string>() {"Артикул", "Заполнения тип", "Заполнения ширина", "Заполнения высота", "Место установки", "Количество всего" },
            OneGroupOneSheet = false
        };



        public struct ItemQuantity { public int Этаж; public int Этажей; } // Количество одинаковых этажей
        public struct Item 
        { 
            public string Плоскость; 
            public string FileName; 
            public string ДетальFileName;
            public string HashProviderFileName;
            public string СтраницаЧертежа;
            public int Стоек;
            public ItemQuantity[] Этажи;
        }

        public static Item Изделия1 = new Item
        {
            Плоскость = плоскость1,
            FileName = "Этаж12_Чертежи.grb",
            HashProviderFileName = "Стойка.grb",
            ДетальFileName = "Деталь_Стойка.grb",
            СтраницаЧертежа = "Чертёж",
            Стоек = 29,
            Этажи = new ItemQuantity[]
            {
            new ItemQuantity {Этаж = 2, Этажей = 1},
            new ItemQuantity {Этаж = 3, Этажей = 25},
            new ItemQuantity {Этаж = 4, Этажей = 25},
            new ItemQuantity {Этаж = 19, Этажей = 2},
            new ItemQuantity {Этаж = 20, Этажей = 1},
            new ItemQuantity {Этаж = 40, Этажей = 1},
            new ItemQuantity {Этаж = 57, Этажей = 1}
            }
        };

        public static Item Изделия2 = new Item
        {
            Плоскость = плоскость2,
            FileName = "Этаж12_Чертежи.grb",
            ДетальFileName = "Деталь_Стойка.grb",
            HashProviderFileName = "Стойка.grb",
            СтраницаЧертежа = "Чертёж",
            Стоек = 29,
            Этажи = new ItemQuantity[]
           {
            new ItemQuantity {Этаж = 2, Этажей = 1},
            new ItemQuantity {Этаж = 3, Этажей = 25},
            new ItemQuantity {Этаж = 4, Этажей = 25},
            new ItemQuantity {Этаж = 19, Этажей = 2},
            new ItemQuantity {Этаж = 20, Этажей = 2},
            new ItemQuantity {Этаж = 57, Этажей = 1}
           }
        };

        public static Item Изделия3 = new Item
        {
            Плоскость = плоскость3,
            FileName = "Изделие3_Чертежи.grb",
            ДетальFileName = "Деталь_Стойка2.grb",
            HashProviderFileName = "Стойка2.grb",
            СтраницаЧертежа = "Чертёж",
            Стоек = 35,
            Этажи = new ItemQuantity[]
            {
            new ItemQuantity {Этаж = 2, Этажей = 14},
            new ItemQuantity {Этаж = 3, Этажей = 12},
            new ItemQuantity {Этаж = 19, Этажей = 1},
            new ItemQuantity {Этаж = 29, Этажей = 13},
            new ItemQuantity {Этаж = 30, Этажей = 14},
            new ItemQuantity {Этаж = 39, Этажей = 1},
            new ItemQuantity {Этаж = 57, Этажей = 1}
            }
        };

        public static Item Изделия4 = new Item
        {
            Плоскость = плоскость4,
            FileName = "Изделие4_Чертежи.grb",
            ДетальFileName = "Деталь_Стойка2.grb",
            HashProviderFileName = "Стойка2.grb",
            СтраницаЧертежа = "Чертёж",
            Стоек = 36,
            Этажи = new ItemQuantity[]
                   {
            new ItemQuantity {Этаж = 2, Этажей = 14},
            new ItemQuantity {Этаж = 3, Этажей = 12},
            new ItemQuantity {Этаж = 19, Этажей = 1},
            new ItemQuantity {Этаж = 29, Этажей = 13},
            new ItemQuantity {Этаж = 30, Этажей = 14},
            new ItemQuantity {Этаж = 39, Этажей = 1},
            new ItemQuantity {Этаж = 57, Этажей = 1}
                   }
        };
    }

    /// <summary>
    /// Класс для нумерации листов.
    /// </summary>
    public class SheetsNumsStore
    {
        public int Начало_Нумерации;
        public int Изделий;

        public int Листов_Чертёж;
        public int Листов_Заказ;
        public int Листов_Отгрузка;

        // Вспомогательные свойства для определения начала каждой секции
        private int Начало_Секции_Заказ => Начало_Нумерации + (Изделий * Листов_Чертёж);
        private int Начало_Секции_Отгрузка => Начало_Секции_Заказ + (Изделий * Листов_Заказ);

        // Общее количество листов
        public int Листов => Начало_Нумерации - 1 + (Изделий * (Листов_Чертёж + Листов_Заказ + Листов_Отгрузка));

        public SheetsNumsStore(int Изделий, int Лист_Начала, int Листов_Чертёж, int Листов_Заказ, int Листов_Отгрузка)
        {
            this.Изделий = Изделий;
            this.Начало_Нумерации = Лист_Начала;
            this.Листов_Чертёж = Листов_Чертёж;
            this.Листов_Заказ = Листов_Заказ;
            this.Листов_Отгрузка = Листов_Отгрузка;
        }

        public int Лист_Чертёж(int Изделие)
        {
            if (Изделие <= 0) return 0;
            // Просто идем подряд: Лист 10, 13, 16...
            return Начало_Нумерации + (Изделие - 1) * Листов_Чертёж;
        }

        public int Лист_Заказ(int Изделие)
        {
            if (Изделие <= 0) return 0;
            // Начинаем после ВСЕХ чертежей
            return Начало_Секции_Заказ + (Изделие - 1) * Листов_Заказ;
        }

        public int Лист_Отгрузка(int Изделие)
        {
            if (Изделие <= 0) return 0;
            // Начинаем после ВСЕХ чертежей и ВСЕХ заказов
            return Начало_Секции_Отгрузка + (Изделие - 1) * Листов_Отгрузка;
        }
    }

    /*
    public class _SheetsNumsStore
    {
        public int Начало_Нумерации;

        public int Листов_Чертёж;
        public int Листов_Заказ;
        public int Листов_Отгрузка;

        public int СекцияПриложение { get { return (Начало_Нумерации + Изделий * Листов_Чертёж); } }

        public int Изделий;

        public int Изделие = 0;
        public int Листов { get { return (Начало_Нумерации - 1 + Изделий*(Листов_Чертёж + Листов_Заказ + Листов_Отгрузка)); } }

        /// <summary>
        /// Ввод параметров для нумерации листов.
        /// Предполагается, что сначала обрабатываются листы "Чертёж", потом листы "Заказ", а потом ""
        /// </summary>
        /// <param name="Изделий">Общее количество изделий</param>
        /// <param name="Лист_Чертёж">Номер Листа, с которого начинается нумерация</param>
        /// <param name="Листов_Чертёж">Количество листов "Чертёж" в выгрузке</param>
        /// <param name="Листов_Заказ">Количество листов "Заказ" в выгрузке</param>
        /// <param name="Листов_Отгрузка">Количество листов "Отгрузка" в выгрузке</param>
        
        public _SheetsNumsStore(int Изделий, int Лист_Чертёж, int Листов_Чертёж, int Листов_Заказ, int Листов_Отгрузка)
        {
            this.Изделий = Изделий;                 // 7
            this.Начало_Нумерации = Лист_Чертёж;    // 10
            this.Листов_Чертёж = Листов_Чертёж;     // 3
            this.Листов_Заказ = Листов_Заказ;       // 2
            this.Листов_Отгрузка = Листов_Отгрузка; //9
        }
        public int Лист_Чертёж(int Изделие)
        {
            if (Изделие == 0) { return (0); }
            return (Начало_Нумерации + (Изделие - 1) * (Листов_Чертёж));
        }
        public int Лист_Заказ(int Изделие)
        {
            if (Изделие == 0) { return (0); }
            return (СекцияПриложение + (Изделие - 1) * (Листов_Заказ + Листов_Отгрузка));
        }
        public int Лист_Отгрузка(int Изделие)
        {
            if (Изделие == 0) { return (0); }
            return (СекцияПриложение + Листов_Заказ + (Изделие - 1) * (Листов_Заказ + Листов_Отгрузка));
        }
    }
    */
}
