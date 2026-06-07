using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniExcelLibs;

namespace SpecCollector
{
    /// <summary>
    /// Класс для чтения файла Этажи.xlsx.
    /// Формат: один лист, строка 1 = заголовки (A=Этаж, B=(5-1)-(5-6), и т.д.),
    ///         дальше строки с данными: колонка A = этаж, остальные = типовой этаж.
    /// Возвращает номер типового этажа по плоскости и этажу.
    /// </summary>
    public class FloorMatcher
    {
        private readonly Dictionary<string, Dictionary<int, int>> _data;

        public string FilePath { get; private set; }

        public FloorMatcher(string activeDocumentPath)
        {
            if (string.IsNullOrEmpty(activeDocumentPath))
                throw new ArgumentException("Путь к активному документу не может быть пустым", nameof(activeDocumentPath));

            string directory = Path.GetDirectoryName(activeDocumentPath);
            if (string.IsNullOrEmpty(directory))
                throw new InvalidOperationException($"Не удалось определить директорию для пути: {activeDocumentPath}");

            FilePath = Path.Combine(directory, "Этажи.xlsx");

            if (!File.Exists(FilePath))
                throw new FileNotFoundException($"Файл 'Этажи.xlsx' не найден: {FilePath}");

            _data = new Dictionary<string, Dictionary<int, int>>(StringComparer.OrdinalIgnoreCase);
            Load();
        }

        private void Load()
        {
            Console.WriteLine($"[FloorMatcher] Loading: {FilePath}");
            
            var sheetNames = MiniExcel.GetSheetNames(FilePath).ToList();
            Console.WriteLine($"[FloorMatcher] Found {sheetNames.Count} sheets: {string.Join(", ", sheetNames)}");
            
            foreach (var sheetName in sheetNames)
            {
                var rows = MiniExcel.Query(FilePath, sheetName: sheetName).ToList();
                Console.WriteLine($"[FloorMatcher] Sheet '{sheetName}': {rows.Count} rows");
                
                if (rows.Count < 2) continue;

                // Первая строка - заголовки: ключи = A, B, C, D, E, значения = "Этаж", "(5-1)-(5-6)"...
                var headerRow = rows[0] as IDictionary<string, object>;
                if (headerRow == null) continue;

                // Строим маппинг: ключ колонки (A/B/C...) -> название плоскости
                var colToPlane = new Dictionary<string, string>();
                foreach (var kvp in headerRow)
                {
                    string colKey = kvp.Key; // "A", "B", "C"...
                    string planeName = kvp.Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(planeName) && !planeName.Equals("Этаж", StringComparison.OrdinalIgnoreCase))
                    {
                        colToPlane[colKey] = planeName;
                    }
                }

                Console.WriteLine($"[FloorMatcher] Column -> Plane mapping: {string.Join(", ", colToPlane.Select(kvp => $"{kvp.Key}={kvp.Value}"))}");

                // Теперь читаем данные: для каждой плоскости строим словарь этаж -> типовой этаж
                foreach (var colPlane in colToPlane)
                {
                    string colKey = colPlane.Key;  // "B", "C"...
                    string planeName = colPlane.Value; // "(5-1)-(5-6)"
                    
                    var sheet = new Dictionary<int, int>();
                    
                    for (int i = 1; i < rows.Count; i++)
                    {
                        var row = rows[i] as IDictionary<string, object>;
                        if (row == null) continue;

                        object floorObj = null;
                        object typicalFloorObj = null;
                        row.TryGetValue("A", out floorObj);           // Колонка A = этаж
                        row.TryGetValue(colKey, out typicalFloorObj); // Колонка B/C/D/E = типовой этаж

                        if (floorObj == null || typicalFloorObj == null) continue;
                        if (!int.TryParse(floorObj.ToString(), out int floor)) continue;
                        if (!int.TryParse(typicalFloorObj.ToString(), out int typicalFloor)) continue;

                        sheet[floor] = typicalFloor;
                    }

                    Console.WriteLine($"[FloorMatcher] Plane '{planeName}': loaded {sheet.Count} floor entries");
                    _data[planeName] = sheet;
                }
            }
        }

        /// <summary>
        /// Возвращает номер типового этажа по плоскости и этажу.
        /// </summary>
        /// <param name="plane">Имя плоскости (например, "(5-1)-(5-6)")</param>
        /// <param name="floor">Этаж</param>
        /// <returns>Номер типового этажа или null, если не найдено</returns>
        public int? GetTypicalFloor(string plane, int floor)
        {
            if (string.IsNullOrEmpty(plane)) return null;

            if (!_data.TryGetValue(plane, out var sheet)) return null;
            if (!sheet.TryGetValue(floor, out var typicalFloor)) return null;

            return typicalFloor;
        }

        /// <summary>
        /// Возвращает статистику по типовым этажам в диапазоне этажей для указанной плоскости.
        /// </summary>
        /// <param name="plane">Название плоскости</param>
        /// <param name="floorFrom">Начальный этаж (включительно)</param>
        /// <param name="floorTo">Конечный этаж (включительно)</param>
        /// <returns>Список пар (типовой этаж, количество вхождений)</returns>
        public List<KeyValuePair<int, int>> GetTypicalFloorStats(string plane, int floorFrom, int floorTo)
        {
            var result = new List<KeyValuePair<int, int>>();

            if (string.IsNullOrEmpty(plane)) return result;
            if (!_data.TryGetValue(plane, out var sheet)) return result;

            var counts = sheet
                .Where(kvp => kvp.Key >= floorFrom && kvp.Key <= floorTo)
                .GroupBy(kvp => kvp.Value)
                .Select(g => new KeyValuePair<int, int>(g.Key, g.Count()))
                .OrderBy(kvp => kvp.Key)
                .ToList();

            return counts;
        }

        /// <summary>
        /// Возвращает количество вхождений указанного типового этажа в диапазоне этажей для заданной плоскости.
        /// </summary>
        /// <param name="plane">Название плоскости</param>
        /// <param name="floorFrom">Начальный этаж (включительно)</param>
        /// <param name="floorTo">Конечный этаж (включительно)</param>
        /// <param name="targetTypicalFloor">Искомый типовой этаж</param>
        /// <returns>Количество вхождений типового этажа в диапазоне</returns>
        public int CountTypicalFloorInRange(string plane, int floorFrom, int floorTo, int targetTypicalFloor)
        {
            if (string.IsNullOrEmpty(plane)) return 0;
            if (!_data.TryGetValue(plane, out var sheet)) return 0;

            return sheet
                .Where(kvp => kvp.Key >= floorFrom && kvp.Key <= floorTo)
                .Count(kvp => kvp.Value == targetTypicalFloor);
        }

        /// <summary>
        /// Возвращает все данные для отладки (плоскость -> этаж -> типовой этаж).
        /// </summary>
        public IReadOnlyDictionary<string, Dictionary<int, int>> GetAllData() => _data;

        public void Reload() { _data.Clear(); Load(); }
    }
}
