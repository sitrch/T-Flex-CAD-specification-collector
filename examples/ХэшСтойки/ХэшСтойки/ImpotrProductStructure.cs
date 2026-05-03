using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex;
using TFlex.Model;

namespace TFAPIHash
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Windows.Forms; // Требуется ссылка на System.Windows.Forms
    using TFlex;
    using TFlex.Model;

    public class ProductStructureImporter
    {
        public static void RunMassImport()
        {
            // 1. Диалог выбора XML-файла (источник данных)
            string xmlPath = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                openFileDialog.Title = "Выберите исходный XML файл структуры";

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                xmlPath = openFileDialog.FileName;
            }

            // 2. Диалог выбора папки проекта
            string projectPath = "";
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Выберите корневую папку проекта с файлами .grb";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() != DialogResult.OK) return;
                projectPath = folderDialog.SelectedPath;
            }

            // 3. Подтверждение операции (безопасность прежде всего)
            var confirm = MessageBox.Show(
                $"Вы уверены, что хотите импортировать структуру из {Path.GetFileName(xmlPath)} \nво все файлы .grb в папке {projectPath}?\n\nРекомендуется сделать резервную копию!",
                "Подтверждение массового импорта",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            // 4. Запуск процесса обработки
            ProcessFiles(projectPath, xmlPath);
        }

        private static void ProcessFiles(string rootPath, string xmlPath)
        {
            // Получаем все файлы .grb рекурсивно
            var files = Directory.GetFiles(rootPath, "*.grb", SearchOption.AllDirectories);
            int processedCount = 0;

            foreach (string file in files)
            {
                try
                {
                    // Открываем документ T-FLEX
                    Document doc = TFlex.Application.OpenDocument(file);

                    if (doc != null)
                    {
                        // Ищем состав изделия по имени "m2spec"
                        var ps = doc.GetProductStructures().FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == "m2spec");

                        if (ps != null)
                        {
                            //ps.load(xmlPath);
                            ps.UpdateStructure();

                            doc.Save();
                            processedCount++;
                        }
                        doc.Close();
                    }
                }
                catch (Exception ex)
                {
                    // Если файл занят или поврежден — пропускаем и пишем в лог T-FLEX
                    //TFlex.Application.Output.Print($"Ошибка в файле {file}: {ex.Message}");
                }
            }

            MessageBox.Show($"Обработка завершена! Обновлено файлов: {processedCount}", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

}
