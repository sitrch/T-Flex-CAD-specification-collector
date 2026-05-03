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

        /// <summary>
        /// Преобразует строку в словарь для записи в Excel через MiniExcel.
        /// Ключи — заголовки столбцов.
        /// </summary>
        public Dictionary<string, object> ToExcelRow()
        {
 return new Dictionary<string, object>
            {
                { "Плоскость", (object)Плоскость ?? DBNull.Value },
                { "Этаж", Этаж.HasValue ? (object)Этаж.Value : DBNull.Value },
                { "Раздел", (object)Раздел ?? DBNull.Value },
                { "Источник", (object)Источник ?? DBNull.Value },
                { "Место установки", (object)МестоУстановки ?? DBNull.Value },
                { "Артикул", (object)Артикул ?? DBNull.Value },
                { "Артикул базовый", (object)АртикулБазовый ?? DBNull.Value },
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
                { "Этажей", Этажей.HasValue ? (object)Этажей.Value : DBNull.Value }
             };
        }

        /// <summary>
        /// Возвращает список имён столбцов в порядке, соответствующем ToExcelRow().
        /// </summary>
        public static List<string> GetColumnNames()
        {
            return new List<string>
            {
                "Плоскость",
                "Этаж",
                "Раздел",
                "Источник",
                "Место установки",
                "Артикул",
                "Артикул базовый",
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
                "Этажей"
             };
        }
    }
}