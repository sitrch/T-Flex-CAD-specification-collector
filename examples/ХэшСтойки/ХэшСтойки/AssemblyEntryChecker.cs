namespace ХэшСтойки
{
    using System;
    using System.IO;
    using TFlex.Model;
    using TFlex.Model.Model2D;

    public class AssemblyEntryChecker
    {
        public static string bom_name = "m2spec";
        public static Document document = TFlex.Application.ActiveDocument;
        static string fname = $"{document.FileName}_Отчёт о включении в спецификацию.txt";
        static string dfpath = document.FilePath;
        public static string filePath = fname; // Путь к файлу

        public void run()
        {
            using (StreamWriter writer = new StreamWriter(filePath, false))
            {
                writer.WriteLine($"Отчет по спецификации '{bom_name}' от {DateTime.Now}");
                writer.WriteLine(new string('-', 50));

                foreach (var child in document.GetFragments())
                {
                    if (child is Fragment subFragment)
                    {
                        CheckRecursive(subFragment, 0, writer);

                    }
                }
            }
        }

        public static void CheckRecursive(Fragment fragment, int level, StreamWriter writer)
        {
            string frname = fragment.Name;
            Document doc = fragment.GetFragmentDocument(true);
            string dname = doc.FileName;
            string indent = new string(' ', level + 1);
            IncludeInBom status = fragment.get_IncludeInSpecificBom(bom_name);

            if (fragment.UseAssemblyStatus == false)
            {
                writer.WriteLine($"changind UseAssemblyStatus to true {fragment.Name} (Файл: {fragment.FilePath})");
                //doc.BeginChanges(""); todo open in context

                //fragment.UseAssemblyStatus = true;
                //doc.EndChanges();
                //doc.Save();
            }

            //if (status != IncludeInBom.WithEmbeddedElements)
            {
                writer.WriteLine($"{indent}[{status}] {fragment.Name} (Файл: {fragment.FilePath})");
                //Console.WriteLine($"!!!{fragment.Name}:{status.ToString()}");
            }
            //Console.WriteLine($"{indent}[{status}] {fragment.Name}");

            // Условие выхода: если статус запрещает вложенные элементы, не идем глубже
            if (status == IncludeInBom.NotIncluded || status == IncludeInBom.WithoutEmbeddedElements)
                return;

            // Перебираем вложенные элементы напрямую через дочерние объекты фрагмента
            foreach (var child in doc.GetFragments())
            {
                CheckRecursive(child, level + 1, writer);
            }
        }
    }

}
