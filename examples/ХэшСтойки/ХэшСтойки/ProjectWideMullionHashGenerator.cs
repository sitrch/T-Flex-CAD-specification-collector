using MiniExcelLibs; // Не забудьте добавить namespace
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using TFlex.Model;
using static TFAPIHash.Data;


namespace TFAPIHash
{
    public class ProjectWideMullionHashGenerator
    {
        Document document = TFlex.Application.ActiveDocument;

        static Document DatabaseDocument;
        string DatabaseDocumentFileName = "dbHash.grb";

        HashDataTable hTable;

        InternalDatabase database;
        public string DatabaseName = "HashStore";

        public string path;
        static string HashProviderPath;
        static string copyDirectoryPath;

        // Список для накопления данных для Excel
        List<object> excelData = new List<object>();


        public ProjectWideMullionHashGenerator()
        {
            path = document.FilePath;
            copyDirectoryPath = Path.Combine(path, Data.copyDirectoryName);
            DirectoryInfo parent = Directory.GetParent(path)?.Parent;
            if (parent != null)
            {
                HashProviderPath = Path.Combine(parent.FullName, "Блоки");
            }
            
            DatabaseDocument = db.FindDBase(document, DatabaseDocumentFileName);
            database = db.GetInternalDatabase(DatabaseDocument, DatabaseName);
            hTable = new HashDataTable(document, database);
        }
        /// <summary>
        /// !Доделать сохранение артикулов стоек и номеров листов в базу данных для публикации в ...
        /// </summary>
        public void СоздатьЧертежиСтоек()
        {
            int Лист = 26;
            int Листов = 190;
            hTable.CheckColumns();
            hTable.ReadRows(database);

            //string excelFolderPath = path;// Папка для Excel
            string excelPath = Path.Combine(path, "Реестр_стоек.xlsx");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            string[] Planes12 = { Data.плоскость1, Data.плоскость2 };
            Лист = СоздатьЧертёжСтойки(Planes12, Data.dname12, Лист, Листов);

            string[] Planes34 = { Data.плоскость3, Data.плоскость4 };
            СоздатьЧертёжСтойки(Planes34, Data.dname34, Лист + 1, Листов);

            
            ExportToExcel(excelData, excelPath);

            ___DebugService.Log($"Файл Excel сохранен: {excelPath}");


        }

        public int СоздатьЧертёжСтойки(string[] Planes, string OriginFileName, int NumerationStart, int Листов)
        {
            ___DebugService.Show();
            ___DebugService.Log($"Старт генерации чертежей стоек.");
            string dpath = Path.Combine(path, OriginFileName);
            Document doc = TFlex.Application.OpenDocument(dpath);

            string pdfFolderPath = Path.Combine(path, Data.ДеталиDirectoryName, Data.PDFDirectoryName);
            string dwgFolderPath = Path.Combine(path, Data.ДеталиDirectoryName, Data.DWGDirectoryName);
           
            Directory.CreateDirectory(pdfFolderPath);
            Directory.CreateDirectory(dwgFolderPath);

            DataTable uniqueTable = hTable.AsEnumerable()
    .Where(row => Planes.Contains(row.Field<string>("Плоскость")))
    .GroupBy(row => row.Field<string>("hash"))
    .Select(g =>
    {
        var row = g.First();
        // Записываем количество найденных дублей в колонку Count
        row.SetField("Count", g.Count());
        return row;
    })
    .OrderBy(row => row.Field<string>("Плоскость"))
    .CopyToDataTable();

            int counter = NumerationStart - 1;
            foreach (DataRow row in uniqueTable.Rows)
            {
                counter++;
                // if (counter > NumerationStart + 10) { break; }
                string Плоскость = row["Плоскость"].ToString();
                int этаж = Convert.ToInt32(row["Этаж"]);
                int стойка = Convert.ToInt32(row["Стойка"]);

                ___DebugService.Log($"Плоскость={Плоскость} Этаж={этаж} Стойка={стойка}");

                doc.BeginChanges("");
                db.SetTextDocumentVariableValue(doc, Data.плоскостьVarName, Плоскость);
                db.SetIntegerDocumentVariableValue(doc, Data.ЭтажVarName, этаж);
                db.SetIntegerDocumentVariableValue(doc, Data.СтойкаVarName, стойка);
                db.SetIntegerDocumentVariableValue(doc, Data.ЛистVarName, counter);
                db.SetIntegerDocumentVariableValue(doc, Data.ЛистовVarName, Листов);

                db.SetIntegerDocumentVariableValue(doc, Data.Стойка_ДетальVarName, 1);

                doc.EndChanges();
                doc.Changed = true;
                doc.Regenerate(new RegenerateOptions { Full = true });

                string АртикулСтойки = db.GetTextDocumentVariableValue(doc,Data.АртикулСтойкиVarName);

                string _name = counter.ToString("D3") + "_" + Плоскость + "-" + этаж.ToString() + "-" + стойка.ToString();

                ExportToDWG(doc, Path.Combine(dwgFolderPath, _name + ".dwg"));
                ExportToPDF(doc, Path.Combine(pdfFolderPath, _name + ".pdf"));

                // Добавляем данные в список для Excel
                excelData.Add(new
                {
                    Лист = counter,
                    Плоскость = Плоскость,
                    Этаж = этаж,
                    АртикулСтойки = АртикулСтойки,
                    Стойка = стойка,
                    Позиция = Convert.ToInt32(row["Позиция"]),
                    Count = Convert.ToInt32(row["Count"]),
                    Листов = Листов,
                    Hash = row["hash"].ToString()

                });
            }
            doc.Close();
            return (counter);
        }

        public void ExportToDWG(Document document, string path)
        {
            ExportToDWG dwgExport = new ExportToDWG(document);
            dwgExport.AutocadExportFileVersion = AutocadExportFileVersionType.efACAD13;
            dwgExport.ConvertToLines = true;
            dwgExport.OrientPagesInModelSpace = false;

            dwgExport.BiarcInterpolationForSplines = false;
            dwgExport.ConvertAreas = true;
            dwgExport.ConvertDimensions = 0;
            dwgExport.ConvertLineText = false;
            dwgExport.ConvertMultitext = 1;

            dwgExport.ExportAllPages = true;
            dwgExport.NotCreateBlocksForPages = false;

            ICollection<Page> ipages = document.GetPages();
            List<Page> pages = new List<Page>();
            foreach (Page p in ipages)
            {
                if (p.Name.Contains("Деталь"))
                {
                    pages.Add(p);
                }
            }

            if (pages.Count == 0) { throw new Exception($"PDF: Страница для экспорта не найдена"); }

            dwgExport.ExportPages = pages;

            dwgExport.Export(path);
        }

        public void ExportToPDF(Document document, string path)
        {
            ExportToPDF pdfExport = new ExportToPDF(document);
            pdfExport.ConvertToRaster = false;
            pdfExport.Export3DModel = false;
            pdfExport.Monochrome = false;

            pdfExport.OpenExportFile = false;

            ICollection<Page> ipages = document.GetPages();
            List<Page> pages = new List<Page>();
            foreach (Page p in ipages)
            {
                if (p.Name.Contains("Деталь"))
                {
                    pages.Add(p);
                }
            }

            if (pages.Count == 0) { throw new Exception($"PDF: Страница для экспорта не найдена"); }

            pdfExport.ExportPages = pages;

            pdfExport.Export(path);
        }

        public void GenerateHashes()
        {
            ___DebugService.Show();
            Directory.CreateDirectory(copyDirectoryPath);
            ___DebugService.Log($"Старт генерации хэшей.");
            GenerateHashesOnPlane(Data.Изделия1);
            GenerateHashesOnPlane(Data.Изделия2);
            GenerateHashesOnPlane(Data.Изделия3);
            GenerateHashesOnPlane(Data.Изделия4);
            ___DebugService.Log($"Окончание генерации хэшей.");

        }
        //document.BeginChanges("");
        /// <summary>
        /// !!! Сборка : Этаж
        /// Плоскости 
        /// (5-1)-(5-6)
        /// (5-6)-(5-1)
        /// Запускает макрос вложенного фрагмента
        /// </summary>
        ///         "nSpace.Hash.HashMacro"
        public void GenerateHashesOnPlane(Item Изделия)
        {
            string fpath = Path.Combine(HashProviderPath, Изделия.HashProviderFileName);
            Document doc = TFlex.Application.OpenDocument(fpath);
            for (int i = 0; i < Изделия.Этажи.Length; i++)
            {
                for (int j = 0; j < Изделия.Стоек;j++)
                {
                    int Стойка = j + 1;
                    ___DebugService.Log($"Плоскость={Изделия.Плоскость} Этаж={Изделия.Этажи[i].Этаж} Стойка={Стойка}");
                    doc.BeginChanges("");
                    db.SetTextDocumentVariableValue(doc, Data.плоскостьVarName, Изделия.Плоскость);
                    db.SetIntegerDocumentVariableValue(doc, Data.ЭтажVarName, Изделия.Этажи[i].Этаж);
                    db.SetIntegerDocumentVariableValue(doc, Data.СтойкаVarName, Стойка);
                    doc.EndChanges();
                    doc.Changed = true;
                    doc.Regenerate(new RegenerateOptions { Full = true });

                    RunMacroInDocument(doc, "nSpace.Hash.WriteHashToDatabaseMacro");

                    //doc.Regenerate(new RegenerateOptions { Full = true });

                    //string fn = Изделия.Плоскость + "-" + Изделия.Изделия[i].Этаж.ToString() + "-" + Изделия.ДетальFileName;
                    //doc.SaveCopy(Path.Combine(copyDirectoryPath, fn));
                    //break; // debug
                }
                //break; // debug
            }

            doc.Close();
        }

        public void RunMacroInDocument(Document doc, string macroFullName)
        {
            // Проверяем наличие макроса в проекте документа
            if (!doc.MacroExists(macroFullName))
            {
                return;
            }

            doc.RunMacro(macroFullName);

        }

        public void ExportToExcel(IEnumerable<object> data, string path)
        {
            // MiniExcel автоматически создаст заголовки на основе имен свойств объектов
            MiniExcel.SaveAs(path, data);
        }

    }
}
