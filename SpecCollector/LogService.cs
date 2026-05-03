using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpecCollector
{
    /// <summary>
    /// Сервис логирования с дедупликацией записей.
    /// Записи накапливаются в памяти, дубликаты подсчитываются.
    /// В конце работы Flush() записывает все записи в файл.
    /// </summary>
    public class LogService
    {
        private class LogEntry
        {
            public string Message;
            public DateTime FirstTime;
            public int Count = 1;
        }

        private readonly string _filePath;
        private readonly DateTime _startTime;
        private readonly Dictionary<string, LogEntry> _entries = new Dictionary<string, LogEntry>();

        /// <param name="directoryPath">Директория головного документа (Document.FilePath)</param>
        public LogService(string directoryPath)
        {
            _filePath = Path.Combine(directoryPath, "collector.log1");
            _startTime = DateTime.Now;

            // Очистить файл и записать шапку
            File.WriteAllText(_filePath, $"[{_startTime}] === Спецификация: старт ===\n");
        }

        /// <summary>
        /// Добавить запись в лог. Если сообщение уже существует — увеличивается счётчик вхождений.
        /// </summary>
        public void Log(string message)
        {
            if (_entries.TryGetValue(message, out var entry))
            {
                entry.Count++;
            }
            else
            {
                _entries[message] = new LogEntry
                {
                    Message = message,
                    FirstTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// Записать все накопленные записи в файл лога.
        /// Файл перезаписывается: шапка + все записи.
        /// </summary>
        public void Flush()
        {
            var lines = new List<string>
            {
                $"[{_startTime}] === Спецификация: старт ==="
            };

            foreach (var entry in _entries.Values.OrderBy(e => e.FirstTime))
            {
                if (entry.Count == 1)
                {
                    lines.Add($"[{entry.FirstTime}] {entry.Message}");
                }
                else
                {
                    lines.Add($"[{entry.FirstTime}] [x{entry.Count}] {entry.Message}");
                }
            }

            lines.Add($"[{DateTime.Now}] === Спецификация: завершение ({_entries.Count} записей) ===");

            File.WriteAllLines(_filePath, lines);
        }
    }
}