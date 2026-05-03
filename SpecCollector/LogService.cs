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
        private readonly Dictionary<string, LogEntry> _entries = new Dictionary<string, LogEntry>();

        public LogService(string filePath)
        {
            _filePath = filePath;
            // Создать пустой файл
            try
            {
                string dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.WriteAllText(filePath, "");
            }
            catch { }
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
        /// </summary>
        public void Flush()
        {
            try
            {
                var lines = new List<string>();

                foreach (var entry in _entries.Values.OrderBy(e => e.FirstTime))
                {
                    if (entry.Count == 1)
                    {
                        lines.Add($"[{entry.FirstTime}] {entry.Message}");
                    }
                    else
                    {
                        lines.Add($"[{entry.FirstTime}] {entry.Message} [x{entry.Count}]");
                    }
                }

                File.WriteAllLines(_filePath, lines);
            }
            catch { }
        }
    }
}