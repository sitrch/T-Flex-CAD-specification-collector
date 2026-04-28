using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpecCollector
{
    /// <summary>
    /// Класс для экспорта собранных записей спецификации в файл Excel.
    /// </summary>
    public class ExcelExporter
    {
        private readonly string _filePath;

        public ExcelExporter(string filePath)
        {
            _filePath = filePath;
        }

        /// <summary>
        /// Экспортирует список записей спецификации в файл Excel.
        /// </summary>
        /// <param name="rows">Список записей спецификации.</param>
        public void Export(List<SpecificationRow> rows)
        {
            // Удаляем существующий файл, если он есть
            if (File.Exists(_filePath))
            {
                File.Delete(_filePath);
            }

            // Создаём директорию, если не существует
            string dir = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // Преобразуем записи в словари для MiniExcel
            var data = new List<Dictionary<string, object>>();
            foreach (var row in rows)
            {
                data.Add(row.ToExcelRow());
            }

            var config = new OpenXmlConfiguration()
            {
                FastMode = true,
                EnableAutoWidth = true
            };

            // Если данных нет — создаём пустой файл с заголовками
            if (data.Count == 0)
            {
                var emptyRow = new Dictionary<string, object>();
                foreach (string col in SpecificationRow.GetColumnNames())
                {
                    emptyRow[col] = DBNull.Value;
                }
                var emptyList = new List<Dictionary<string, object>> { emptyRow };
                MiniExcel.SaveAs(_filePath, emptyList, overwriteFile: true, configuration: config);
            }
            else
            {
                MiniExcel.SaveAs(_filePath, data, overwriteFile: true, configuration: config);
            }
        }
    }
}