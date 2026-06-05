using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniExcelLibs;
using TFlex;

namespace SpecCollector
{
    /// <summary>
    /// Класс для чтения файла Захватки.xlsx (фаз).
    /// Строка = Этаж, Столбец = Mullion (стойка).
    /// </summary>
    public class PhaseReader
    {
        private readonly Dictionary<string, Dictionary<int, Dictionary<string, object>>> _data;

        public string FilePath { get; private set; }

        public PhaseReader(string activeDocumentPath)
        {
            if (string.IsNullOrEmpty(activeDocumentPath))
                throw new ArgumentException("Путь к активному документу не может быть пустым", nameof(activeDocumentPath));

            string directory = Path.GetDirectoryName(activeDocumentPath);
            if (string.IsNullOrEmpty(directory))
                throw new InvalidOperationException($"Не удалось определить директорию для пути: {activeDocumentPath}");

            FilePath = Path.Combine(directory, "Захватки.xlsx");

            if (!File.Exists(FilePath))
                throw new FileNotFoundException($"Файл 'Захватки.xlsx' не найден: {FilePath}");

            _data = new Dictionary<string, Dictionary<int, Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase);
            Load();
        }

        private void Load()
        {
            foreach (var sheetName in MiniExcel.GetSheetNames(FilePath))
            {
                var sheet = new Dictionary<int, Dictionary<string, object>>();
                var rows = MiniExcel.Query(FilePath, sheetName: sheetName).ToList();

                if (rows.Count == 0)
                {
                    _data[sheetName] = sheet;
                    continue;
                }

                // Header row: Этаж, м1, м2, м3...
                var header = rows[0] as IDictionary<string, object>;
                if (header == null) continue;

                // Mullion columns start with "м" (м1, м2...). Case-insensitive.
                var mullionCols = header.Keys
                    .Where(k => k != null && k.ToString().StartsWith("м", StringComparison.OrdinalIgnoreCase))
                    .Select(k => k.ToString())
                    .ToList();

                // Data rows
                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i] as IDictionary<string, object>;
                    if (row == null) continue;

                    object floorObj = null;
                    row.TryGetValue("Этаж", out floorObj);
                    if (floorObj == null) continue;
                    if (!int.TryParse(floorObj.ToString(), out int floor)) continue;

                    var floorData = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (var mullion in mullionCols)
                    {
                        object val = null;
                        row.TryGetValue(mullion, out val);
                        floorData[mullion] = val;
                    }
                    sheet[floor] = floorData;
                }

                _data[sheetName] = sheet;
            }
        }

        /// <summary>
        /// Возвращает значение ячейки на пересечении этажа (строка) и моллиона (столбец).
        /// </summary>
        /// <param name="phase">Имя фазы (лист книги)</param>
        /// <param name="mullion">Моллион/стойка, например "м7" или "М7"</param>
        /// <param name="floor">Этаж (значение в столбце "Этаж")</param>
        /// <returns>Значение ячейки или null</returns>
        public object GetValue(string phase, string mullion, int floor)
        {
            if (string.IsNullOrEmpty(phase)) throw new ArgumentException(nameof(phase));
            if (string.IsNullOrEmpty(mullion)) throw new ArgumentException(nameof(mullion));

            var key = mullion.ToLowerInvariant();

            if (!_data.TryGetValue(phase, out var sheet)) return null;
            if (!sheet.TryGetValue(floor, out var row)) return null;
            if (!row.TryGetValue(key, out var value)) return null;

            return value;
        }

        public string GetString(string phase, string mullion, int floor) => GetValue(phase, mullion, floor)?.ToString();

        public double? GetDouble(string phase, string mullion, int floor)
        {
            var v = GetValue(phase, mullion, floor);
            if (v == null) return null;
            if (v is double d) return d;
            if (v is int i) return i;
            if (v is float f) return f;
            if (decimal.TryParse(v.ToString(), out var dec)) return (double)dec;
            if (double.TryParse(v.ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var res)) return res;
            return null;
        }

        public int? GetInt(string phase, string mullion, int floor)
        {
            var d = GetDouble(phase, mullion, floor);
            return d.HasValue ? (int)d.Value : (int?)null;
        }

        public void Reload() { _data.Clear(); Load(); }
    }
}