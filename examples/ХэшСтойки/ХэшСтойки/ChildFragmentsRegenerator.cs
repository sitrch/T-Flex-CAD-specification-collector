using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;
using TFlex.Model;
using TFlex.Model.Model2D;
using System.IO;


namespace TFAPIHash
{
    public class ChildFragmentsRegenerator
    {
        public Document document = TFlex.Application.ActiveDocument;

        public ChildFragmentsRegenerator()
        {
            ___DebugService.Show();
        }

        // 1. Вызов через диалог выбора файла
        public void RegenerateChildFragments()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Документы T-FLEX (*.grb)|*.grb|Все файлы (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    RegenerateChildFragments(ofd.FileName);
                }
            }
        }

        public void RegenerateChildFragments(Document mainDoc)
        {
            ___DebugService.Log($"Начало обработки сборки: {mainDoc.FileName}");

            // Запуск рекурсии
            ProcessRecursiveWithContext(mainDoc, 0);

            // Финальное сохранение корня после всех вложенных обновлений
            mainDoc.Regenerate(new RegenerateOptions { Full = true });
            mainDoc.Save();

            ___DebugService.Log("--- Готово! Вся иерархия обновлена сверху вниз ---");

        }
        public void RegenerateChildFragments(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                ___DebugService.Log($"[ОШИБКА] Файл не найден: {filePath}");
                return;
            }

            
            ___DebugService.Log($"--- Запуск обработки: {Path.GetFileName(filePath)} ---");

            // 2. Открываем корневой документ
            using (Document mainDoc = TFlex.Application.OpenDocument(filePath, false))
            {
                if (mainDoc == null)
                {
                    ___DebugService.Log("Ошибка: Не удалось открыть выбранный файл.");
                    return;
                }

                ___DebugService.Log($"Начало обработки сборки: {Path.GetFileName(filePath)}");

                // Запуск рекурсии
                ProcessRecursiveWithContext(mainDoc, 0);

                // Финальное сохранение корня после всех вложенных обновлений
                mainDoc.Regenerate(new RegenerateOptions { Full = true });
                mainDoc.Save();

                ___DebugService.Log("--- Готово! Вся иерархия обновлена сверху вниз ---");
                //___DebugService.Close();
            }
        }

        private void ProcessRecursiveWithContext(Document currentDoc, int level)
        {
            var fragments = currentDoc.GetFragments();
            string indent = new string(' ', level * 2);

            foreach (Fragment frag in fragments)
            {
                if (string.IsNullOrEmpty(frag.FilePath)) continue;

                // 1. Регенерируем фрагмент в контексте родителя
                frag.Regenerate(true);

                // 2. Получаем документ фрагмента
                Document subDoc = frag.GetFragmentDocument(true);

                if (subDoc == null) continue;

                // КРИТИЧЕСКИ ВАЖНО: Сразу сохраняем имя в строку, 
                // пока объект subDoc гарантированно "жив"
                string subDocPath = subDoc.FileName;
                string subDocName = Path.GetFileName(subDocPath);

                ___DebugService.Log($"{indent}→ Уровень {level}: {subDocName}");

                try
                {
                    // 3. РЕКУРСИЯ: Идем глубже
                    // Передаем subDoc напрямую. 
                    ProcessRecursiveWithContext(subDoc, level + 1);

                    // 4. Сохранение изменений
                    if (!subDoc.IsReadOnly)
                    {
                       //subDoc.BeginChanges("Обновление из сборки");
                        //subDoc.EndChanges();
                        //subDoc.Save();
                    }
                }
                finally
                {
                    // Вместо using используем явное закрытие, 
                    // чтобы точно контролировать момент уничтожения объекта
                    //subDoc.Close();
                }
            }

        }

        public void UpdateFragmentsInContext()
        {
            ___DebugService.Show();
            // Получаем все 3D фрагменты сборки
            var fragments = document.GetFragments();
            int count = fragments.Count;
            int counter = 0;
            foreach (Fragment frag in fragments)
            {
                counter++;
                ___DebugService.Log($" {counter} / {count} ");

                frag.Regenerate(true);
                Document d = frag.GetFragmentDocument();
                d.Save();
            }
        }



    }
}
