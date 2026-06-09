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
    /// Использует плоский словарь с кортежным ключом (плоскость, этаж) -> типовой этаж.
    /// </summary>
    public class FloorMatcher
    {
        private readonly Dictionary<(string plane, int floor), int> _data;

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

            _data = new Dictionary<(string, int), int>();
            Load();
        }

        private void Load()
        {
            var sheetNames = MiniExcel.GetSheetNames(FilePath).ToList();

            foreach (var sheetName in sheetNames)
            {
                var rows = MiniExcel.Query(FilePath, sheetName: sheetName).ToList();
                if (rows.Count < 2) continue;

                var headerRow = rows[0] as IDictionary<string, object>;
                if (headerRow == null) continue;

                var colToPlane = new Dictionary<string, string>();
                foreach (var kvp in headerRow)
                {
                    string colKey = kvp.Key;
                    string planeName = kvp.Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(planeName) && !planeName.Equals("Этаж", StringComparison.OrdinalIgnoreCase))
                    {
                        colToPlane[colKey] = planeName;
                    }
                }

                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i] as IDictionary<string, object>;
                    if (row == null) continue;

                    object floorObj = null;
                    row.TryGetValue("A", out floorObj);
                    if (floorObj == null) continue;
                    if (!int.TryParse(floorObj.ToString(), out int floor)) continue;

                    foreach (var colPlane in colToPlane)
                    {
                        string colKey = colPlane.Key;
                        string planeName = colPlane.Value;

                        object typicalFloorObj = null;
                        row.TryGetValue(colKey, out typicalFloorObj);
                        if (typicalFloorObj == null) continue;
                        if (!int.TryParse(typicalFloorObj.ToString(), out int typicalFloor)) continue;

                        var key = (planeName, floor);
                        _data[key] = typicalFloor;
                    }
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

            var key = (plane, floor);
            return _data.TryGetValue(key, out var typicalFloor) ? typicalFloor : (int?)null;
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

            var counts = _data
                .Where(kvp => kvp.Key.plane.Equals(plane, StringComparison.OrdinalIgnoreCase))
                .Where(kvp => kvp.Key.floor >= floorFrom && kvp.Key.floor <= floorTo)
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

            return _data
                .Where(kvp => kvp.Key.plane.Equals(plane, StringComparison.OrdinalIgnoreCase))
                .Where(kvp => kvp.Key.floor >= floorFrom && kvp.Key.floor <= floorTo)
                .Count(kvp => kvp.Value == targetTypicalFloor);
        }

        /// <summary>
        /// Возвращает все данные для отладки.
        /// </summary>
        public IReadOnlyDictionary<(string plane, int floor), int> GetAllData() => _data;


        /// <summary>
        /// Возвращает список уникальных типовых этажей для указанной плоскости.
        /// </summary>
        /// <param name="plane">Название плоскости</param>
        /// <returns>Отсортированный список уникальных типовых этажей</returns>
        public List<int> GetUniqueFloorList(string plane)
        {
            if (string.IsNullOrEmpty(plane)) return new List<int>();

            return _data
                .Where(kvp => kvp.Key.plane.Equals(plane, StringComparison.OrdinalIgnoreCase))
                .Select(kvp => kvp.Value)
                .Distinct()
                .OrderBy(x => x)
                .ToList();
        }
        public void Reload() { _data.Clear(); Load(); }
    }
}
