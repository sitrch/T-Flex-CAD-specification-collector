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

        /// <summary>
        /// Экспортирует список записей спецификации в файл Excel.
        /// </summary>
        /// <param name="rows">Список записей спецификации.</param>
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
                        // Создаём листы по разделам
                        var sheets = new Dictionary<string, object>();

                        // Главный лист со всеми данными
                        sheets.Add("Спецификация", data);

                        // Группируем данные по разделам
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

                        string[] sheetOrder = { "Детали", "Термомосты", "Заполнения", "Комплектующие", "Материалы", "Раскрой", "Проверка", "Другое" };

                        foreach (var sheetName in sheetOrder)
                        {
                            if (groups.ContainsKey(sheetName))
                            {
                                sheets.Add(sheetName, groups[sheetName]);
                            }
                        }

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
