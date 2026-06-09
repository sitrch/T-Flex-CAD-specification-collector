using System;
using System.Collections.Generic;

namespace SpecCollector
{
    /// <summary>
    /// Модель данных одной записи спецификации, извлечённой из фрагмента.
    /// Столбцы соответствуют столбцам спецификации T-FLEX.
    /// </summary>
    public class SpecificationRow
    {
        /// <summary>Источник записи — имя файла документа (фрагмента), из которого получена запись.</summary>
        public string Источник { get; set; }

        /// <summary>Плоскость — из переменной $Плоскость корневого документа.</summary>
        public string Плоскость { get; set; }

        /// <summary>Этаж — из переменной $Этаж корневого документа.</summary>
        public int? Этаж { get; set; }

        /// <summary>Номер стойки (моллион) из переменной Стойка фрагментов стоек</summary>
        public int? MullionNumber { get; set; }

        /// <summary>Этажей — количество одинаковых этажей из SpecItemQuantity.</summary>
        public int? Этажей { get; set; }

        /// <summary>Артикул</summary>
        public string Артикул { get; set; }

        /// <summary>Артикул базовый</summary>
        public string АртикулБазовый { get; set; }

        /// <summary>Длина</summary>
        public double? Длина { get; set; }

        /// <summary>Длина, м</summary>
        public double? ДлинаМ { get; set; }

        /// <summary>Наименование</summary>
        public string Наименование { get; set; }

        /// <summary>Значение всего</summary>
        public double? ЗначениеВсего { get; set; }

        /// <summary>Раздел</summary>
        public string Раздел { get; set; }

        /// <summary>Количество всего</summary>
        public int? КоличествоВсего { get; set; }

        /// <summary>Заполнения тип</summary>
        public string ЗаполненияТип { get; set; }

        /// <summary>Заполнения ширина</summary>
        public double? ЗаполненияШирина { get; set; }

        /// <summary>Заполнения высота</summary>
        public double? ЗаполненияВысота { get; set; }

        /// <summary>Заполнения толщина</summary>
        public double? ЗаполненияТолщина { get; set; }

        /// <summary>Заполнения площадь</summary>
        public double? ЗаполненияПлощадь { get; set; }

        /// <summary>Место установки</summary>
        public string МестоУстановки { get; set; }

        /// <summary>Размещение (Ригель/Стойка)</summary>
        public string Размещение { get; set; }

        /// <summary>Этап (захватка) из PhaseReader</summary>
        public string Этап { get; set; }

        /// <summary>Код заказа</summary>
        public string КодЗаказа { get; set; }

        /// <summary>Динамические столбцы на основе RequestItems (Name -> значение floorCount)</summary>
        public Dictionary<string, int> DynamicColumns { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// Преобразует строку в словарь для записи в Excel через MiniExcel.
        /// Ключи — заголовки столбцов.
        /// Стойка сразу после Этаж.
        /// Количество_{Name} вычисляется в C#:
        ///   — для Материалы: Длина,м × Количество всего × Заказ_{Name}
        ///   — для остальных: Количество всего × Заказ_{Name}
        /// </summary>
        public Dictionary<string, object> ToExcelRow()
        {
            var dict = new Dictionary<string, object>();

            dict["Плоскость"] = (object)Плоскость ?? DBNull.Value;
            dict["Этаж"] = Этаж.HasValue ? (object)Этаж.Value : DBNull.Value;
            dict["Стойка"] = MullionNumber.HasValue ? (object)MullionNumber.Value : DBNull.Value;
            dict["Раздел"] = (object)Раздел ?? DBNull.Value;
            dict["Источник"] = (object)Источник ?? DBNull.Value;
            dict["Место установки"] = (object)МестоУстановки ?? DBNull.Value;
            dict["Артикул"] = (object)Артикул ?? DBNull.Value;
            dict["Артикул базовый"] = (object)АртикулБазовый ?? DBNull.Value;
            dict["Код заказа"] = (object)КодЗаказа ?? DBNull.Value;
            dict["Наименование"] = (object)Наименование ?? DBNull.Value;
            dict["Длина"] = Длина.HasValue ? (object)Длина.Value : DBNull.Value;
            dict["Длина, м"] = ДлинаМ.HasValue ? (object)ДлинаМ.Value : DBNull.Value;
            dict["Значение всего"] = ЗначениеВсего.HasValue ? (object)ЗначениеВсего.Value : DBNull.Value;
            dict["Количество всего"] = КоличествоВсего.HasValue ? (object)КоличествоВсего.Value : DBNull.Value;
            dict["Заполнения тип"] = (object)ЗаполненияТип ?? DBNull.Value;
            dict["Заполнения ширина"] = ЗаполненияШирина.HasValue ? (object)ЗаполненияШирина.Value : DBNull.Value;
            dict["Заполнения высота"] = ЗаполненияВысота.HasValue ? (object)ЗаполненияВысота.Value : DBNull.Value;
            dict["Заполнения толщина"] = ЗаполненияТолщина.HasValue ? (object)ЗаполненияТолщина.Value : DBNull.Value;
            dict["Заполнения площадь"] = ЗаполненияПлощадь.HasValue ? (object)ЗаполненияПлощадь.Value : DBNull.Value;
            dict["Этажей"] = Этажей.HasValue ? (object)Этажей.Value : DBNull.Value;
            dict["Размещение"] = (object)Размещение ?? DBNull.Value;

            // Определяем, относится ли строка к разделу Материалы
            bool isMaterial = Раздел != null
                && SpecData.BOMSections.TryGetValue(Раздел, out var group)
                && group == "Материалы";

            if (DynamicColumns != null)
            {
                foreach (var kvp in DynamicColumns)
                {
                    dict["Заказ_" + kvp.Key] = kvp.Value;

                    if (isMaterial && ДлинаМ.HasValue)
                    {
                        dict["Количество_" + kvp.Key] = Math.Round(ДлинаМ.Value * (КоличествоВсего ?? 0) * kvp.Value, 2);
                    }
                    else
                    {
                        int? кол = КоличествоВсего * kvp.Value;
                        dict["Количество_" + kvp.Key] = кол.HasValue ? (object)кол.Value : DBNull.Value;
                    }
                }
            }

            return dict;
        }

        /// <summary>
        /// Преобразует строку в словарь для экспорта в Спецификации_Этапы.xlsx.
        /// Без Этажей, с колонкой Этап.
        /// Всего = Количество всего (для Материалы: Длина,м × Количество всего).
        /// </summary>
        public Dictionary<string, object> ToExcelRowStages()
        {
            var dict = new Dictionary<string, object>
            {
                { "Плоскость", (object)Плоскость ?? DBNull.Value },
                { "Этаж", Этаж.HasValue ? (object)Этаж.Value : DBNull.Value },
                { "Стойка", MullionNumber.HasValue ? (object)MullionNumber.Value : DBNull.Value },
                { "Раздел", (object)Раздел ?? DBNull.Value },
                { "Источник", (object)Источник ?? DBNull.Value },
                { "Место установки", (object)МестоУстановки ?? DBNull.Value },
                { "Артикул", (object)Артикул ?? DBNull.Value },
                { "Артикул базовый", (object)АртикулБазовый ?? DBNull.Value },
                { "Код заказа", (object)КодЗаказа ?? DBNull.Value },
                { "Наименование", (object)Наименование ?? DBNull.Value },
                { "Длина", Длина.HasValue ? (object)Длина.Value : DBNull.Value },
                { "Длина, м", ДлинаМ.HasValue ? (object)ДлинаМ.Value : DBNull.Value },
                { "Значение всего", ЗначениеВсего.HasValue ? (object)ЗначениеВсего.Value : DBNull.Value },
                { "Количество всего", КоличествоВсего.HasValue ? (object)КоличествоВсего.Value : DBNull.Value },
                { "Заполнения тип", (object)ЗаполненияТип ?? DBNull.Value },
                { "Заполнения ширина", ЗаполненияШирина.HasValue ? (object)ЗаполненияШирина.Value : DBNull.Value },
                { "Заполнения высота", ЗаполненияВысота.HasValue ? (object)ЗаполненияВысота.Value : DBNull.Value },
                { "Заполнения толщина", ЗаполненияТолщина.HasValue ? (object)ЗаполненияТолщина.Value : DBNull.Value },
                { "Заполнения площадь", ЗаполненияПлощадь.HasValue ? (object)ЗаполненияПлощадь.Value : DBNull.Value },
                { "Размещение", (object)Размещение ?? DBNull.Value },
                { "Этап", (object)Этап ?? DBNull.Value }
            };

            bool isMaterial = Раздел != null
                && SpecData.BOMSections.TryGetValue(Раздел, out var group)
                && group == "Материалы";

            if (isMaterial && ДлинаМ.HasValue)
            {
                dict["Всего"] = Math.Round(ДлинаМ.Value * (КоличествоВсего ?? 0), 2);
            }
            else
            {
                dict["Всего"] = КоличествоВсего.HasValue ? (object)КоличествоВсего.Value : DBNull.Value;
            }

            return dict;
        }

        /// <summary>
        /// Возвращает список имён столбцов в порядке, соответствующем ToExcelRow().
        /// Включает динамические столбцы из RequestItems с префиксом Заказ_.
        /// </summary>
        public static List<string> GetColumnNames()
        {
            var columns = new List<string>
            {
                "Плоскость",
                "Этаж",
                "Стойка",
                "Раздел",
                "Источник",
                "Место установки",
                "Артикул",
                "Артикул базовый",
                "Код заказа",
                "Наименование",
                "Длина",
                "Длина, м",
                "Значение всего",
                "Количество всего",
                "Заполнения тип",
                "Заполнения ширина",
                "Заполнения высота",
                "Заполнения толщина",
                "Заполнения площадь",
                "Этажей",
                "Размещение"
            };

            if (SpecData.RequestItems != null)
            {
                foreach (var item in SpecData.RequestItems)
                {
                    if (!string.IsNullOrEmpty(item.Name))
                    {
                        columns.Add("Заказ_" + item.Name);
                        columns.Add("Количество_" + item.Name);
                    }
                }
            }

            return columns;
        }
    }
}