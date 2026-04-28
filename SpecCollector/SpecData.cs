using System.Collections.Generic;

namespace SpecCollector
{
    public static class SpecData
    {
        public static string ПлоскостиDirectoryName = "Плоскости";
        
        public static string плоскостьVarName = "$Плоскость";
        public static string ЭтажVarName = "Этаж";
        
        public static string плоскость1 = "(5-1)-(5-6)";
        public static string плоскость2 = "(5-6)-(5-1)";
        public static string плоскость3 = "(5-А)-(5-Л)";
        public static string плоскость4 = "(5-Л)-(5-А)";

        public struct SpecItemQuantity
        {
            public int Этаж;
            public int Этажей;
        }

        public struct SpecItem
        {
            public string Плоскость;
            public string FileName;
            public SpecItemQuantity[] Этажи;
        }

        public static SpecItem Изделия1 = new SpecItem
        {
            Плоскость = плоскость1,
            FileName = "Этаж12_Чертежи.grb",
            Этажи = new SpecItemQuantity[]
            {
                new SpecItemQuantity { Этаж = 2, Этажей = 1 },
                new SpecItemQuantity { Этаж = 3, Этажей = 25 },
                new SpecItemQuantity { Этаж = 4, Этажей = 25 },
                new SpecItemQuantity { Этаж = 19, Этажей = 2 },
                new SpecItemQuantity { Этаж = 20, Этажей = 1 },
                new SpecItemQuantity { Этаж = 40, Этажей = 1 },
                new SpecItemQuantity { Этаж = 57, Этажей = 1 }
            }
        };

        public static SpecItem Изделия2 = new SpecItem
        {
            Плоскость = плоскость2,
            FileName = "Этаж12_Чертежи.grb",
            Этажи = new SpecItemQuantity[]
            {
                new SpecItemQuantity { Этаж = 2, Этажей = 1 },
                new SpecItemQuantity { Этаж = 3, Этажей = 25 },
                new SpecItemQuantity { Этаж = 4, Этажей = 25 },
                new SpecItemQuantity { Этаж = 19, Этажей = 2 },
                new SpecItemQuantity { Этаж = 20, Этажей = 2 },
                new SpecItemQuantity { Этаж = 57, Этажей = 1 }
            }
        };

        public static SpecItem Изделия3 = new SpecItem
        {
            Плоскость = плоскость3,
            FileName = "Изделие3_Чертежи.grb",
            Этажи = new SpecItemQuantity[]
            {
                new SpecItemQuantity { Этаж = 2, Этажей = 14 },
                new SpecItemQuantity { Этаж = 3, Этажей = 12 },
                new SpecItemQuantity { Этаж = 19, Этажей = 1 },
                new SpecItemQuantity { Этаж = 29, Этажей = 13 },
                new SpecItemQuantity { Этаж = 30, Этажей = 14 },
                new SpecItemQuantity { Этаж = 39, Этажей = 1 },
                new SpecItemQuantity { Этаж = 57, Этажей = 1 }
            }
        };

        public static SpecItem Изделия4 = new SpecItem
        {
            Плоскость = плоскость4,
            FileName = "Изделие4_Чертежи.grb",
            Этажи = new SpecItemQuantity[]
            {
                new SpecItemQuantity { Этаж = 2, Этажей = 14 },
                new SpecItemQuantity { Этаж = 3, Этажей = 12 },
                new SpecItemQuantity { Этаж = 19, Этажей = 1 },
                new SpecItemQuantity { Этаж = 29, Этажей = 13 },
                new SpecItemQuantity { Этаж = 30, Этажей = 14 },
                new SpecItemQuantity { Этаж = 39, Этажей = 1 },
                new SpecItemQuantity { Этаж = 57, Этажей = 1 }
            }
        };
    }
}