using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SpecCollector
{
    /// <summary>
    /// Класс для экспорта собранных записей спецификации в файл Excel.
    /// </summary>
    public class ExcelExporter
    {
        private string _filePath;

        public ExcelExporter(string filePath)
        {
            _filePath = filePath;
        }

        /// <summary>
        /// Проверяет, занят ли файл другим приложением.
        /// </summary>
        private bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
