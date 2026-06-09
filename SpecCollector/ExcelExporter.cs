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
        /// Проверяет, можно ли создать/перезаписать файл по указанному пути.
        /// Для существующего файла проверяет, не заблокирован ли он.
        /// Для несуществующего — пробует создать и удалить временный файл.
        /// </summary>
        public static bool ValidateFileCanBeWritten(string fullPath, string displayName)
        {
            if (File.Exists(fullPath))
            {
                if (IsFileLocked(fullPath))
                {
                    ___DisplayService.Log($"Файл занят: {displayName}");
                    return false;
                }
                ___DisplayService.Log($"Файл доступен: {displayName}");
                return true;
            }

            // Файла нет — проверяем, можем ли создать
            try
            {
                string dir = Path.GetDirectoryName(fullPath);
                if (!Directory.Exists(dir))
                {
                    ___DisplayService.Log($"Директория не существует: {dir}");
                    return false;
                }

                string tempFile = Path.Combine(dir, Path.GetRandomFileName());
                using (var fs = File.Create(tempFile)) { }
                File.Delete(tempFile);
                ___DisplayService.Log($"Файл будет создан: {displayName}");
                return true;
            }
            catch (Exception ex)
            {
                ___DisplayService.Log($"Невозможно создать {displayName}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Проверяет доступность входных файлов проекта.
        /// </summary>
        /// <param name="directory">Директория проекта</param>
        /// <returns>True если все входные файлы доступны, иначе false</returns>
        public static bool ValidateInputFiles(string directory)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                return false;

            var errors = new List<string>();

            string floorsFilePath = Path.Combine(directory, SpecData.FloorsFileName);
            string phasesFilePath = Path.Combine(directory, SpecData.PhasesFileName);

            if (!File.Exists(floorsFilePath))
                errors.Add($"Входной файл не найден: {SpecData.FloorsFileName}");
            else if (!CanReadExcelFile(floorsFilePath))
                errors.Add($"Входной файл недоступен для чтения: {SpecData.FloorsFileName}");
            else
                ___DisplayService.Log($"Входной файл доступен: {SpecData.FloorsFileName}");

            if (!File.Exists(phasesFilePath))
                errors.Add($"Входной файл не найден: {SpecData.PhasesFileName}");
            else if (!CanReadExcelFile(phasesFilePath))
                errors.Add($"Входной файл недоступен для чтения: {SpecData.PhasesFileName}");
            else
                ___DisplayService.Log($"Входной файл доступен: {SpecData.PhasesFileName}");

            foreach (var err in errors)
                ___DisplayService.Log(err);

            return errors.Count == 0;
        }

        /// <summary>
        /// Пытается прочитать Excel-файл через MiniExcel для проверки доступности.
        /// </summary>
        public static bool CanReadExcelFile(string filePath)
        {
            try
            {
                var sheetNames = MiniExcel.GetSheetNames(filePath).ToList();
                return sheetNames.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Проверяет, занят ли файл другим приложением.
        /// </summary>
        public static bool IsFileLocked(string filePath)
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

        /// <summary>
        /// Открывает файл с монопольным доступом и возвращает поток.
        /// Поток нужно держать открытым до конца работы программы.
        /// </summary>
        public static FileStream LockFile(string fullPath)
        {
            if (!File.Exists(fullPath))
                return null;

            try
            {
                return new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Экспортирует список записей спецификации в файл Excel.
        /// </summary>
        /// <param name="rows">Список записей спецификации.</param>
        private static string[] SheetOrder = { "Детали", "Термомосты", "Заполнения", "Комплектующие", "Материалы", "Раскрой", "Проверка", "Другое" };

        public static Dictionary<string, object> GroupDataBySections(List<Dictionary<string, object>> data)
        {
            var sheets = new Dictionary<string, object>();
            sheets.Add("Спецификация", data);

            var groupMapping = SpecData.BOMSections;
            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var row in data)
            {
                string razdel = row.ContainsKey("Раздел") ? row["Раздел"]?.ToString() : null;
                string groupName = "Другое";

                if (razdel != null && groupMapping.ContainsKey(razdel))
                {
                    groupName = groupMapping[razdel];
                }

                if (!groups.ContainsKey(groupName))
                {
                    groups[groupName] = new List<Dictionary<string, object>>();
                }
                groups[groupName].Add(row);
            }

            foreach (var sheetName in SheetOrder)
            {
                if (groups.ContainsKey(sheetName))
                {
                    sheets.Add(sheetName, groups[sheetName]);
                }
            }

            return sheets;
        }

        public void Export(List<SpecificationRow> rows)
        {
            int maxAttempts = 3;
            int attempt = 0;

            while (attempt < maxAttempts)
            {
                attempt++;

                // Проверяем, занят ли файл другим приложением
                if (File.Exists(_filePath) && IsFileLocked(_filePath))
                {
                    if (attempt < maxAttempts)
                    {
                        var dialog = new FileLockedDialog(_filePath);
                        dialog.ShowDialog();

                        switch (dialog.Result)
                        {
                            case FileLockedDialog.FileAction.Retry:
                                continue;
                            case FileLockedDialog.FileAction.SaveAs:
                                var saveDialog = new SaveFileDialog
                                {
                                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                                    Title = "Сохранить как",
                                    FileName = Path.GetFileName(_filePath),
                                    InitialDirectory = Path.GetDirectoryName(_filePath)
                                };
                                if (saveDialog.ShowDialog() == DialogResult.OK)
                                {
                                    _filePath = saveDialog.FileName;
                                    continue;
                                }
                                return;
                            case FileLockedDialog.FileAction.Cancel:
                                return;
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Не удалось сохранить файл после {maxAttempts} попыток.\nФайл занят другим приложением:\n{_filePath}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                try
                {
                    // Удаляем существующий файл, если он есть
                    if (File.Exists(_filePath))
                    {
                        File.Delete(_filePath);
                    }

                    // Создаём директорию, если не существует
                    //string dir = Path.GetDirectoryName(_filePath);
                    //if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    //{
                    //Directory.CreateDirectory(dir);
                    //}

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
                        var sheets = GroupDataBySections(data);
                        MiniExcel.SaveAs(_filePath, sheets, overwriteFile: true, configuration: config);
                    }

                    return;
                }
                catch (IOException ex)
                {
                    if (attempt < maxAttempts)
                    {
                        var dialog = new FileLockedDialog(_filePath);
                        dialog.ShowDialog();

                        switch (dialog.Result)
                        {
                            case FileLockedDialog.FileAction.Retry:
                                continue;
                            case FileLockedDialog.FileAction.SaveAs:
                                var saveDialog = new SaveFileDialog
                                {
                                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                                    Title = "Сохранить как",
                                    FileName = Path.GetFileName(_filePath),
                                    InitialDirectory = Path.GetDirectoryName(_filePath)
                                };
                                if (saveDialog.ShowDialog() == DialogResult.OK)
                                {
                                    _filePath = saveDialog.FileName;
                                    continue;
                                }
                                return;
                            case FileLockedDialog.FileAction.Cancel:
                                return;
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Ошибка при сохранении файла:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }
  
    public Dictionary<string, List<Dictionary<string, object>>> FilterBySections(
        List<Dictionary<string, object>> data,
        Dictionary<string, string> BOMSections)
    {
        var result = new Dictionary<string, List<Dictionary<string, object>>>();
        
        // Initialize all group names from BOMSections values
        foreach (var groupName in BOMSections.Values.Distinct())
        {
            result[groupName] = new List<Dictionary<string, object>>();
        }
        
        // Create lookup structures
        var bomSectionKeys = new HashSet<string>(BOMSections.Keys);
        var sectionToGroup = BOMSections.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        
        var unassignedRows = new List<Dictionary<string, object>>();
        
        foreach (var row in data)
        {
            // Get Раздел value
            if (!row.TryGetValue("Раздел", out var sectionObj) || sectionObj == null)
            {
                unassignedRows.Add(row);
                continue;
            }
            
            var sectionStr = sectionObj.ToString();
            
            if (bomSectionKeys.Contains(sectionStr))
            {
                var groupName = sectionToGroup[sectionStr];
                result[groupName].Add(row);
            }
            else
            {
                unassignedRows.Add(row);
            }
        }
        
        // Add unassigned rows if any
        if (unassignedRows.Any())
        {
            result["Нераспределённые"] = unassignedRows;
        }
        
        return result;
    }
    }
}
