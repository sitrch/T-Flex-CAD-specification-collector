using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using MiniExcelLibs;

namespace SpecCollector
{
    public class PhaseReader
    {
        private readonly Dictionary<(string plane, int floor, string mullion), string> _data;

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

            _data = new Dictionary<(string, int, string), string>();
            Load();
        }

        private void Load()
        {
            foreach (var sheetName in MiniExcel.GetSheetNames(FilePath))
            {
                var rows = MiniExcel.Query(FilePath, sheetName: sheetName).ToList();
                if (rows.Count == 0) continue;

                var header = rows[0] as IDictionary<string, object>;
                if (header == null) continue;

                var mullionCols = header.Keys
                    .Where(k => k != null)
                    .Select(k => k.ToString())
                    .Where(s => s.Length > 1 && (s[0] == 'с' || s[0] == 'С') && s.Substring(1).All(char.IsDigit))
                    .ToList();

                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i] as IDictionary<string, object>;
                    if (row == null) continue;

                    if (!row.TryGetValue("Этаж", out var floorObj) || floorObj == null) continue;
                    if (!int.TryParse(floorObj.ToString(), out int floor)) continue;

                    foreach (var mullion in mullionCols)
                    {
                        row.TryGetValue(mullion, out var val);
                        var key = (sheetName, floor, mullion.ToLowerInvariant());
                        _data[key] = val?.ToString();
                    }
                }
            }
        }

        public string GetPhase(string plane, string mullion, int floor)
        {
            if (string.IsNullOrEmpty(plane) || string.IsNullOrEmpty(mullion)) return null;
            var key = (plane, floor, mullion.ToLowerInvariant());
            return _data.TryGetValue(key, out var value) ? value : null;
        }

        public void Reload() { _data.Clear(); Load(); }
    }
}