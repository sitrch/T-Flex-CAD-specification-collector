using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using TFlex;

namespace SpecCollector
{
    /// <summary>
    /// Класс для работы с файлом Захватки.xlsx.
    /// При инициализации открывает файл в папке активного документа и загружает данные в память.
    /// </summary>
    public class ZahvatkiReader
    {
        private readonly Dictionary<string, Dictionary<int, Dictionary<string, object>>> _sheetsData;

        /// <summary>
        /// Путь к файлу Захватки.xlsx
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Список имен листов (плоскостей)
        /// </summary>
        public IReadOnlyList<string> SheetNames => _sheetsData.Keys.ToList().AsReadOnly();

        /// <summary>
        /// Инициализирует reader, открывая файл Захватки.xlsx в папке активного документа.
        /// </summary>
        /// <param name="activeDocumentPath">Путь к папке активного документа (doc.FilePath)</param>
        public ZahvatkiReader(string activeDocumentPath)
        {
            if (string.IsNullOrEmpty(activeDocumentPath))
                throw new ArgumentException("Путь к активному документу не может быть пустым", nameof(activeDocumentPath));

            // Формируем путь к файлу Захватки.xlsx в папке с активным документом
            string directory = Path.GetDirectoryName(activeDocumentPath);
            if (string.IsNullOrEmpty(directory))
                throw new InvalidOperationException($"Не удалось определить директорию для пути: {activeDocumentPath}");

            FilePath = Path.Combine(directory, "Захватки.xlsx");

            if (!File.Exists(FilePath))
                throw new FileNotFoundException($"Файл 'Захватки.xlsx' не найден по пути: {FilePath}");

            _sheetsData = new Dictionary<string, Dictionary<int, Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase);
            LoadData();
        }

        /// <summary>
        /// Загружает все листы файла в память.
        /// Структура: SheetName -> Floor -> RackColumn -> Value
        /// </summary>
        private void LoadData()
        {
            foreach (var sheetName in MiniExcel.GetSheetNames(FilePath))
            {
                var sheetDict = new Dictionary<int, Dictionary<string, object>>();

                // Читаем строки листа. Заголовок: Этаж, с1, с2, с3...
                var rows = MiniExcel.Query(FilePath, sheetName: sheetName).ToList();

                if (rows.Count == 0)
                {
                    _sheetsData[sheetName] = sheetDict;
                    continue;
                }

                // Первая строка - заголовки. Получаем имена колонок стойк.
                var headerRow = rows[0] as IDictionary<string, object>;
                if (headerRow == null)
                    continue;

                // Колонки стойк начинаются со 2й (индекс 1), первая - "Этаж"
                var rackColumns = headerRow.Keys.Where(k => k != null && k.ToString().StartsWith("с", StringComparison.OrdinalIgnoreCase))
                    .Select(k => k.ToString())
                    .ToList();

                // Обрабатываем строки данных (начиная со 2й)
                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i] as IDictionary<string, object>;
                    if (row == null)
                        continue;

                    // Получаем значение этажа
                    object floorObj = null;
                    row.TryGetValue("Этаж", out floorObj);
                    if (floorObj == null)
                        continue;

                    if (!int.TryParse(floorObj.ToString(), out int floor))
                        continue;

                    // Создаем словарь значений для этого этажа
                    var floorData = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

                    foreach (var rackCol in rackColumns)
                    {
                        object value = null;
                        row.TryGetValue(rackCol, out value);
                        floorData[rackCol] = value; // Может быть null
                    }

                    sheetDict[floor] = floorData;
                }

                _sheetsData[sheetName] = sheetDict;
            }
        }

        /// <summary>
        /// Получает значение ячейки на пересечении стойки и этажа для заданной плоскости (листа).
        /// </summary>
        /// <param name="plane">Имя плоскости (соответствует имени листа в книге)</param>
        /// <param name="rack">Стойка в формате "с7" или "C7" (колонка в Excel)</param>
        /// <param name="floor">Этаж (значение в столбце "Этаж")</param>
        /// <returns>Значение ячейки или null, если не найдено</returns>
        public object GetCellValue(string plane, string rack, int floor)
        {
            if (string.IsNullOrEmpty(plane))
                throw new ArgumentException("Плоскость не может быть пустой", nameof(plane));

            if (string.IsNullOrEmpty(rack))
                throw new ArgumentException("Стойка не может быть пустой", nameof(rack));

            // Нормализуем имя стойки к нижнему регистру для сравнения (с1, с2...)
            string normalizedRack = rack.ToLowerInvariant();

            if (!_sheetsData.TryGetValue(plane, out var sheetData))
                return null; // Лист не найден

            if (!sheetData.TryGetValue(floor, out var floorData))
                return null; // Этаж не найден

            if (!floorData.TryGetValue(normalizedRack, out var value))
                return null; // Стойка не найдена

            return value;
        }

        /// <summary>
        /// Получает значение ячейки как строку.
        /// </summary>
        public string GetCellValueAsString(string plane, string rack, int floor)
        {
            var value = GetCellValue(plane, rack, floor);
            return value?.ToString();
        }

        /// <summary>
        /// Получает значение ячейки как число (double).
        /// </summary>
        public double? GetCellValueAsDouble(string plane, string rack, int floor)
        {
            var value = GetCellValue(plane, rack, floor);
            if (value == null)
                return null;

            if (value is double d)
                return d;

            if (value is int i)
                return i;

            if (value is float f)
                return f;

            if (decimal.TryParse(value.ToString(), out decimal dec))
                return (double)dec;

            if (double.TryParse(value.ToString().Replace(',', '.'), 
                System.Globalization.NumberStyles.Any, 
                System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;

            return null;
        }

        /// <summary>
        /// Получает значение ячейки как целое число.
        /// </summary>
        public int? GetCellValueAsInt(string plane, string rack, int floor)
        {
            var value = GetCellValueAsDouble(plane, rack, floor);
            return value.HasValue ? (int?)value.Value : null;
        }

        /// <summary>
        /// Проверяет, существует ли плоскость (лист) в файле.
        /// </summary>
        public bool HasPlane(string plane)
        {
            return _sheetsData.ContainsKey(plane);
        }

        /// <summary>
        /// Проверяет, существует ли этаж в указанной плоскости.
        /// </summary>
        public bool HasFloor(string plane, int floor)
        {
            return _sheetsData.TryGetValue(plane, out var sheetData) && sheetData.ContainsKey(floor);
        }

        /// <summary>
        /// Получает все доступные стойки для заданной плоскости и этажа.
        /// </summary>
        public IEnumerable<string> GetAvailableRacks(string plane, int floor)
        {
            if (_sheetsData.TryGetValue(plane, out var sheetData) &&
                sheetData.TryGetValue(floor, out var floorData))
            {
                return floorData.Keys;
            }
            return Enumerable.Empty<string>();
        }

        /// <summary>
        /// Получает все доступные этажи для заданной плоскости.
        /// </summary>
        public IEnumerable<int> GetAvailableFloors(string plane)
        {
            if (_sheetsData.TryGetValue(plane, out var sheetData))
            {
                return sheetData.Keys.OrderBy(x => x);
            }
            return Enumerable.Empty<int>();
        }

        /// <summary>
        /// Перезагружает данные из файла (если файл изменился).
        /// </summary>
        public void Reload()
        {
            _sheetsData.Clear();
            LoadData();
        }
    }
}