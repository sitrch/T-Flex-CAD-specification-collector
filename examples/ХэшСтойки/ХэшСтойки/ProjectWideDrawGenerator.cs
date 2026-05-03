using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using TFlex.Configuration;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;
using TFlex.Model.Model3D;
using static TFAPIHash.Data;

namespace TFAPIHash
{
    public class ProjectWideDrawGenerator
    {
        Document document = TFlex.Application.ActiveDocument;

        static Document DatabaseDocument;
        //string DatabaseDocumentFileName = "dbHash.grb";

        //HashDataTable hTable;

        //InternalDatabase database;
        //public string DatabaseName = "HashStore";

        public string path;
        static string copyDirectoryPath;

        ExcelMiniDataBlock dblРаскрой = null;
        ExcelMiniDataBlock dblПокупныеИзделия = null;
        ExcelMiniDataBlock dblМатериалы = null;
        ExcelMiniDataBlock dblЗаполнения = null;


        public bool OnlyBOMGenerate = false;
        public string fn_ProductStructure;




        public ProjectWideDrawGenerator()
        {
            path = document.FilePath;
            copyDirectoryPath = Path.Combine(path, Data.copyDirectoryName);


            //List<ExcelMiniDataBlock> dblЧетежи;


            //DatabaseDocument = db.FindDBase(document, DatabaseDocumentFileName);
            //database = db.GetInternalDatabase(DatabaseDocument, DatabaseName);
            //hTable = new HashDataTable(document, database);
        }

        public void GenerateDraw()
        {
            OnlyBOMGenerate = true;

            SheetsNumsStore Листы1 = new SheetsNumsStore(7, 20, 3, 3, 2);
            SheetsNumsStore Листы2 = new SheetsNumsStore(6, 20, 3, 3, 2);
            SheetsNumsStore Листы3 = new SheetsNumsStore(7, 20, 3, 3, 2);
            SheetsNumsStore Листы4 = new SheetsNumsStore(7, 20, 3, 3, 2);

            ___DebugService.Show();

            Directory.CreateDirectory(Path.Combine(path, Data.ПлоскостиDirectoryName));

            // Файл для записи спецификаций
            fn_ProductStructure = Path.Combine(path, Data.ПлоскостиDirectoryName, "Спецификации.xlsx");
            ExcelMiniDataBlock block = new ExcelMiniDataBlock(fn_ProductStructure, "");
            block.CreateEmpty();// создаём пустой файл

            GenerateDraw(Data.Изделия1, Листы1);
            GenerateDraw(Data.Изделия2, Листы2);
            GenerateDraw(Data.Изделия3, Листы3);
            GenerateDraw(Data.Изделия4, Листы4);

            dblРаскрой.Write();
            dblПокупныеИзделия.Write();
            dblМатериалы.Write();
            dblЗаполнения.Write();
        }

        public void GenerateDraw(Item Изделия, SheetsNumsStore Листы)
        {
            ___DebugService.Log($"Старт генерации: {Изделия.Плоскость}");

            SheetsNumsStore result = Листы;

            string dpath = Path.Combine(path, Изделия.FileName);
            Document doc = TFlex.Application.OpenDocument(dpath);
            ___DebugService.Log($"  Документ загружен: {doc.FileName}");

            string pdfFolderPath = Path.Combine(path, Data.ПлоскостиDirectoryName, Изделия.Плоскость, Data.PDFDirectoryName);
            string dwgFolderPath = Path.Combine(path, Data.ПлоскостиDirectoryName, Изделия.Плоскость, Data.DWGDirectoryName);
            string xlsxFolderPath = Path.Combine(path, Data.ПлоскостиDirectoryName, Изделия.Плоскость);
            Directory.CreateDirectory(pdfFolderPath);
            Directory.CreateDirectory(dwgFolderPath);

            

            // Файл для записи названий чертежей
            string fn_Чертежи = Path.Combine(xlsxFolderPath, Изделия.Плоскость + "_Чертежи.xlsx");
            ExcelMiniDataBlock DataBlockЧертежи = new ExcelMiniDataBlock(fn_Чертежи, "Чертежи");
            DataBlockЧертежи.CreateEmpty();// создаём пустой файл

            int counter_страниц = 0;
            int counter_изделий = 0;

            // цикл по изделиям
            foreach (ItemQuantity item in Изделия.Этажи)
            {
                counter_изделий++;

                counter_страниц = Листы.Лист_Чертёж(counter_изделий);
                // цикл по списку способов отбражения чертежа
                foreach (string v in Data.Видимость)
                {
                    string ТипЧертежа = $"Этаж {item.Этаж}. {v}.";
                    doc.BeginChanges("");
                    db.SetTextDocumentVariableValue(doc, Data.плоскостьVarName, Изделия.Плоскость);
                    db.SetIntegerDocumentVariableValue(doc, Data.ЭтажVarName, item.Этаж);
                    db.SetTextDocumentVariableValue(doc, Data.ВидимостьVarName, v);

                    db.SetIntegerDocumentVariableValue(doc, Data.Лист_ЧертёжVarName, counter_страниц);
                    db.SetIntegerDocumentVariableValue(doc, Data.Лист_ЗаказVarName, result.Лист_Заказ(counter_изделий));
                    db.SetIntegerDocumentVariableValue(doc, Data.Лист_ОтгрузкаVarName, result.Лист_Отгрузка(counter_изделий));
                    db.SetIntegerDocumentVariableValue(doc, Data.ЛистовVarName, Листы.Листов);

                    db.SetIntegerDocumentVariableValue(doc, Data.Стойка_ДетальVarName, 0); // Настраиваем отобраение 

                    doc.EndChanges();
                    doc.Changed = true;

                    ___DebugService.Log($"  Регенерация с новыми переменными: ");
                    ___DebugService.Log($"     Плоскость = {Изделия.Плоскость} ");
                    ___DebugService.Log($"     Этаж = {item.Этаж} ");
                    ___DebugService.Log($"     Видимость = {v} ");

                    ChildFragmentsRegenerator chgen = new ChildFragmentsRegenerator();
                    chgen.RegenerateChildFragments(doc);

                    doc.BeginChanges("");
                    /*
                    doc.Regenerate(new RegenerateOptions { Full = true,
                        NotifyPlugins = true,
                        Projections = true,
                        UpdateAllLinks = true,
                        UpdateBillOfMaterials = true,
                        UpdateDrawingViews = true,
                        UpdateProductStructures = true,
                        UpdateSymbols = true
                    });
                    */


                    UpdateProductStructure(Data.СоставИзделия);

                    doc.EndChanges();

                    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
                    if (OnlyBOMGenerate)
                    {
                        break;
                    }

                    SaveAll(doc, DataBlockЧертежи, counter_страниц, pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, Изделия.СтраницаЧертежа, v);
                    counter_страниц++;   
                }

                //if (!OnlyBOMGenerate)
                {
                    SaveAll(doc, DataBlockЧертежи, result.Лист_Заказ(counter_изделий), pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, "Профили (раскрой)", "Профили (раскрой)");
                    SaveAll(doc, DataBlockЧертежи, result.Лист_Заказ(counter_изделий) + 1, pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, "Покупные изделия (заказ)", "Покупные изделия (заказ)");
                    SaveAll(doc, DataBlockЧертежи, result.Лист_Заказ(counter_изделий) + 2, pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, "Заполнения (заказ)", "Заполнения (заказ)");

                    SaveAll(doc, DataBlockЧертежи, result.Лист_Отгрузка(counter_изделий), pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, "Детали (отгрузка)", "Детали (отгрузка)");
                    SaveAll(doc, DataBlockЧертежи, result.Лист_Отгрузка(counter_изделий)+1, pdfFolderPath, dwgFolderPath, Изделия.Плоскость, item.Этаж, "Комплектующие (отгрузка)", "Комплектующие (отгрузка)");
                }
                
                ProductStructureExporter pse = new ProductStructureExporter(doc,fn_ProductStructure);

                ___DebugService.Log($"  Exporting product structure... ");

                //List<Dictionary<string, object>> DataList;

                Dictionary<string, object> ColumnsPlusDict = new Dictionary<string, object>() {
                    { "Плоскость", Изделия.Плоскость },
                    { "Этаж",item.Этаж },
                    { "Количество изделий",item.Этажей },
                };

                pse.ExportProductStructure(Data.Раскрой, ref dblРаскрой, ColumnsPlusDict);
                pse.ExportProductStructure(Data.ПокупныеИзделия, ref dblПокупныеИзделия, ColumnsPlusDict);
                pse.ExportProductStructure(Data.Материалы, ref dblМатериалы, ColumnsPlusDict);
                pse.ExportProductStructure(Data.Заполнения, ref dblЗаполнения, ColumnsPlusDict);

                //break;

            }
            doc.Close();

            if (!OnlyBOMGenerate)
            {
                DataBlockЧертежи.Write();
            }
        }
        /// <summary>
        /// Выполняет комплексное сохранение чертежа: экспорт в DWG и PDF, 
        /// а также регистрацию данных о листе в блоке для Excel.
        /// </summary>
        /// <param name="doc">Документ, из которого производится экспорт.</param>
        /// <param name="Чертежи">Объект накопления данных для последующей выгрузки в Excel.</param>
        /// <param name="counter">Порядковый номер листа (используется для индексации имен файлов).</param>
        /// <param name="pdfFolderPath">Путь к папке для сохранения PDF-файлов.</param>
        /// <param name="dwgFolderPath">Путь к папке для сохранения DWG-файлов.</param>
        /// <param name="Плоскость">Маркировка плоскости (часть имени файла).</param>
        /// <param name="Этаж">Номер этажа (используется в имени и описании).</param>
        /// <param name="СтраницаЧертежа">Идентификатор или имя конкретного листа/вида для печати.</param>
        /// <param name="ВидЧертежа">Тип содержимого (например, "Схема" или "План").</param>
        public void SaveAll(Document doc, ExcelMiniDataBlock Чертежи, int counter, 
                                    string pdfFolderPath, string dwgFolderPath, 
                                    string Плоскость, int Этаж, 
                                    string СтраницаЧертежа, string ВидЧертежа)
        {
            string fn = $"{counter.ToString("D3")}_{Плоскость}-{Этаж.ToString()}_{ВидЧертежа}";
            string ТипЧертежа = $"Этаж {Этаж}. {ВидЧертежа}.";

            ___DebugService.Log($"  Saving dwg... ");
            ExportToDWG(doc, СтраницаЧертежа, Path.Combine(dwgFolderPath, fn + ".dwg"));

            ___DebugService.Log($"  Saving pdf... ");
            ExportToPDF(doc, СтраницаЧертежа, Path.Combine(pdfFolderPath, fn + ".pdf"));

            var excelRow = new Dictionary<string, object>()
            {
                ["Лист"] = counter,
                ["Название"] = ТипЧертежа,
                ["Примечание"] = ""
            };
            Чертежи.AddRow(excelRow);
           //dblЧетежи.Add(Чертежи);
        } 
        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="SheetName"></param>
        /// <param name="path"></param>
        /// <exception cref="Exception"></exception>
        public void ExportToDWG(Document document, string SheetName, string path)
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
                if (p.Name.Contains(SheetName))
                {
                    pages.Add(p);
                }
            }

            if (pages.Count == 0) { throw new Exception($"PDF: Страница для экспорта не найдена"); }

            dwgExport.ExportPages = pages;

            dwgExport.Export(path);
        }

        public void ExportToPDF(Document document, string SheetName, string path)
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
                if (p.Name.Contains(SheetName))
                {
                    pages.Add(p);
                }
            }

            if (pages.Count == 0) { throw new Exception($"PDF: Страница для экспорта не найдена"); }

            pdfExport.ExportPages = pages;

            pdfExport.Export(path);
        }

        public void UpdateProductStructure(string СоставИзделия)
        {
            ProductStructure ps = document.GetProductStructures()
                 .FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == СоставИзделия);
            if (ps == null)
            {
                throw new Exception($"Состав изделия {СоставИзделия} не найден!");
            }

            document.BeginChanges("");

            ps.UpdateStructure();
            ps.UpdateReports();

            document.EndChanges();
        }
    }

    
}
