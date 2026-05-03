using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFAPIHash;
using TFlex.Command;
using TFlex.Model;

namespace TFAPIHash
{
    public class RewriteProductStructure
    {

        public void ProcessDocumentsWithSpecificType()
        {
            //string targetTypeName = "Спецификация"; // Имя состава, который нужно создать

            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() != DialogResult.OK) return;

                var files = Directory.GetFiles(fbd.SelectedPath, "*.grb", SearchOption.AllDirectories);

                int fcount = files.Length;
                int counter = 0;

                foreach (var file in files)
                {
                    counter++;
                    ___DebugService.Log($"Обработка {Path.GetFileName(file)} ({counter}/{fcount}).");
                    // Проверка на ReadOnly
                    FileInfo fi = new FileInfo(file);
                    if (fi.IsReadOnly) continue;

                    Document doc = TFlex.Application.OpenDocument(file, false);
                    if (doc == null) continue;

                    ___DebugService.Log($"  Обновление составов изделия... ");
                    UpdateProductStructure(doc, Data.СоставИзделия);
                }
            }
        }
        public void UpdateProductStructure(Document document, string СоставИзделия)
        {
            ProductStructure ps = document.GetProductStructures()
                 .FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == СоставИзделия);
            if (ps == null)
            {
                return;
                //throw new Exception($"Состав изделия {СоставИзделия} не найден!");
            }

            document.BeginChanges("");

            document.Regenerate(new RegenerateOptions { Full = true });

            ps.MarkChanged();
            ps.UpdateStructure();
            ps.UpdateReports();

            document.Regenerate(new RegenerateOptions { Full = true });

            document.EndChanges();

            document.Save();

        }
    }
}
