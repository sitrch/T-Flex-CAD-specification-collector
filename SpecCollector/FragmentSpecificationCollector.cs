using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;
using TFlex.Model.Model2D;
using static TFlex.Model.RowElementGroup;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;

namespace SpecCollector
{
    public class FragmentSpecificationCollector
    {
        private const string ProductStructureName = "m2spec";
        private const string _logPath = @"C:\Users\Lerik\YandexDisk\templates\!Programming\Git\Sobiratel\SpecCollector\bin\Debug\spec_collector.log";
        private readonly List<SpecificationRow> _results = new List<SpecificationRow>();
        private readonly HashSet<string> _processedDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private string _rootПлоскость;
        private int? _rootЭтаж;
        private int? _rootЭтажей;

        public bool OnlyBOMGenerate = false;
        private string _path;

        private void SafeWriteLog(string message, bool isFirstWrite = false)
        {
            try
            {
                var logDir = Path.GetDirectoryName(_logPath);
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                if (isFirstWrite)
                {
                    File.WriteAllText(_logPath, message);
                }
                else
                {
                    File.AppendAllText(_logPath, message);
                }
            }
            catch
            {
                // ignore logging errors
            }
        }

        public void GenerateSpec()
        {
            Document doc = TFlex.Application.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("No active T-FLEX document.");
            }

            OnlyBOMGenerate = true;
            _path = doc.FilePath;

            SpecDebugService.Show();
            SpecDebugService.Log("Start collecting specifications...");
            SpecDebugService.Log("Folder: " + _path);

            string fn_AllSpecs = Path.Combine(_path, "VseSpecifikacii.xlsx");

            var allData = new List<Dictionary<string, object>>();

            CollectSpecData(SpecData.Изделия1, allData);
            CollectSpecData(SpecData.Изделия2, allData);
            CollectSpecData(SpecData.Изделия3, allData);
            CollectSpecData(SpecData.Изделия4, allData);

            ExportToMultipleSheets(fn_AllSpecs, allData);
            SpecDebugService.Log("Done! File: " + fn_AllSpecs);
        }

        private void CollectSpecData(SpecData.SpecItem изделие, List<Dictionary<string, object>> allData)
        {
            string dpath = Path.Combine(_path, изделие.Плоскость) == null ? изделие.FileName : Path.Combine(_path, изделие.FileName);
            SpecDebugService.Log("Обработка: " + изделие.FileName + " (" + изделие.Плоскость + ")");

            Document doc = TFlex.Application.OpenDocument(dpath);
            if (doc == null)
            {
                SpecDebugService.Log("  Ошибка: документ не найден");
                return;
            }

            foreach (var etazh in изделие.Этажи)
            {
                SpecDebugService.Log("  Этаж " + etazh.Этаж + "...");

                doc.BeginChanges("");
                SetTextVariableValue(doc, SpecData.плоскостьVarName, изделие.Плоскость);
                SetIntegerVariableValue(doc, SpecData.ЭтажVarName, etazh.Этаж);
                doc.EndChanges();
                doc.Changed = true;

                UpdateProductStructure(doc);
                CollectFromDocument(doc, изделие.Плоскость, etazh.Этаж, etazh.Этажей);

                foreach (var row in _results)
                {
                    allData.Add(row.ToExcelRow());
                }
                _results.Clear();
            }

            doc.Close();
        }

        private void CollectFromDocument(Document doc, string плоскость, int этаж, int этажей)
        {
            _rootПлоскость = плоскость;
            _rootЭтаж = этаж;
            _rootЭтажей = этажей;
            _processedDocs.Clear();
            ProcessDocument(doc, 0);
        }

        private void ExportToMultipleSheets(string filePath, List<Dictionary<string, object>> allData)
        {
            var sheets = new Dictionary<string, object>();

            sheets.Add("Спецификация", allData);

            var groupMapping = new Dictionary<string, string>
            {
                { @"Детали\Крышка ригеля", "Детали" },
                { @"Детали\Крышка стойки", "Детали" },
                { @"Детали\Прижим ригеля", "Детали" },
                { @"Детали\Прижим стойки", "Детали" },
                { @"Детали\Ригели", "Детали" },
                { @"Детали\Стойки", "Детали" },
                { "Термомосты", "Термомосты" },
                { @"Заполнения", "Заполнения" },
                { @"Заполнения\Изделия", "Заполнения" },
                { @"Комплектующие", "Комплектующие" },
                { @"Комплектующие\Кронштейны", "Комплектующие" },
                { "Уплотнители", "Материалы" },
                { "Листовые материалы", "Материалы" },
                { "Материалы", "Материалы" },
                { "Раскрой", "Раскрой" },
                { "Раскрой деталей", "Раскрой" },
                { @"Проверка\Набор ригеля", "Проверка" },
                { @"Проверка\Набор стойки", "Проверка" }
            };

            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var row in allData)
            {
                string razdel = row.ContainsKey("Razdel") ? row["Razdel"]?.ToString() : null;
                string groupName = "Drugoe";

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

            var config = new OpenXmlConfiguration()
            {
                FastMode = true,
                EnableAutoWidth = true
            };

            MiniExcel.SaveAs(filePath, sheets, overwriteFile: true, configuration: config);
        }

        private void UpdateProductStructure(Document doc)
        {
            var productStructures = doc.GetProductStructures();
            var ps = productStructures?.FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == ProductStructureName);
            if (ps != null)
            {
                ps.UpdateStructure();
            }
        }

        private void SetTextVariableValue(Document doc, string name, string value)
        {
            Variable v = doc.FindVariable(name);
            if (v != null) v.TextValue = value;
        }

        private void SetIntegerVariableValue(Document doc, string name, int value)
        {
            Variable v = doc.FindVariable(name);
            if (v != null) v.RealValue = value;
        }

        public void CollectAndExport()
        {
            Document doc = TFlex.Application.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("No active T-FLEX document.");
            }
            string outputExcelPath = Path.Combine(doc.FilePath, "SpecCollector.xlsx");
            CollectAndExport(doc, outputExcelPath);
        }

        public void CollectAndExport(Document rootDocument, string outputExcelPath)
        {
            _results.Clear();
            _processedDocs.Clear();
            _rootЭтажей = null;

            _rootПлоскость = GetTextVariableValue(rootDocument, SpecData.плоскостьVarName);
            _rootЭтаж = GetIntVariableValue(rootDocument, SpecData.ЭтажVarName);

            SafeWriteLog("=== Started " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ===\r\n", isFirstWrite: true);
            SafeWriteLog("Root: Плоскость=" + _rootПлоскость + ", Этаж=" + _rootЭтаж + "\r\n");

            ProcessDocument(rootDocument, 0);
            var exporter = new ExcelExporter(outputExcelPath);
            exporter.Export(_results);
            SafeWriteLog("=== Finished: " + _results.Count + " rows collected ===\r\n");
        }

        public List<SpecificationRow> Collect(Document rootDocument)
        {
            _results.Clear();
            _processedDocs.Clear();
            _rootЭтажей = null;
            ProcessDocument(rootDocument, 0);
            return new List<SpecificationRow>(_results);
        }

        private void ProcessDocument(Document doc, int level)
        {
            if (doc == null) return;

            string docPath = doc.FileName ?? "";
            if (!string.IsNullOrEmpty(docPath) && _processedDocs.Contains(docPath))
            {
                return;
            }
            if (!string.IsNullOrEmpty(docPath))
            {
                _processedDocs.Add(docPath);
            }

            string docName = string.IsNullOrEmpty(docPath) ? "(bezymyannyj)" : Path.GetFileName(docPath);

            ExtractOwnRecordsFromBOM(doc, docName, level);

            var fragments = doc.GetFragments();
            if (fragments != null)
            {
                foreach (Fragment frag in fragments)
                {
                    if (string.IsNullOrEmpty(frag.FilePath)) continue;

                    try
                    {
                        frag.Regenerate(true);
                        Document subDoc = frag.GetFragmentDocument(true);
                        if (subDoc != null)
                        {
                            ProcessDocument(subDoc, level + 1);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("[Sobiratel] Error processing fragment " + frag.FilePath + ": " + ex.Message);
                    }
                }
            }
        }

        private void ExtractOwnRecordsFromBOM(Document doc, string docName, int level)
        {
            var productStructures = doc.GetProductStructures();
            if (productStructures == null || productStructures.Count == 0)
            {
                SafeWriteLog("  -> NO ProductStructures found\r\n");
                return;
            }

            foreach (var ps in productStructures)
            {
                SafeWriteLog("  -> PS: " + ps?.GetName(ModelObjectName.Name) + "\r\n");
            }

            var productStructure = productStructures?.FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == ProductStructureName);
            if (productStructure == null) return;

            var scheme = productStructure.GetScheme();
            var allRows = productStructure.GetAllRowElements();

            SafeWriteLog("  -> GetAllRowElements: " + (allRows?.Count ?? 0) + " rows\r\n");
            if (allRows != null)
            {
                foreach (var row in allRows)
                {
                    var specRow = ExtractRowFromRowElement(row, scheme, docName);
                    _results.Add(specRow);
                }
            }
        }

        private SpecificationRow ExtractRowFromRowElement(RowElement rowElement, Scheme scheme, string docName)
        {
            var row = new SpecificationRow
            {
                Плоскость = _rootПлоскость,
                Этаж = _rootЭтаж,
                Этажей = _rootЭтажей,
                Источник = docName
            };

            row.Артикул = GetCellValueAsString(rowElement, scheme, "Артикул");
            row.АртикулБазовый = GetCellValueAsString(rowElement, scheme, "Артикул базовый");
            row.Длина = GetCellValueAsDouble(rowElement, scheme, "Длина");
            row.ДлинаМ = GetCellValueAsDouble(rowElement, scheme, "Длина, м");
            row.Наименование = GetCellValueAsString(rowElement, scheme, "Наименование");
            row.ЗначениеВсего = GetCellValueAsDouble(rowElement, scheme, "Значение всего");
            row.Раздел = GetCellValueAsString(rowElement, scheme, "Раздел");
            row.КоличествоВсего = GetCellValueAsInt(rowElement, scheme, "Количество всего");
            row.ЗаполненияТип = GetCellValueAsString(rowElement, scheme, "Заполнения тип");
            row.ЗаполненияШирина = GetCellValueAsDouble(rowElement, scheme, "Заполнения ширина");
            row.ЗаполненияВысота = GetCellValueAsDouble(rowElement, scheme, "Заполнения высота");
            row.ЗаполненияТолщина = GetCellValueAsDouble(rowElement, scheme, "Заполнения толщина");
            row.ЗаполненияПлощадь = GetCellValueAsDouble(rowElement, scheme, "Заполнения площадь");
            var incDoc = rowElement.IncludeInDoc?.Value;
            row.Спецификация = (incDoc != null && (bool)incDoc) ? "1" : "0";
            row.МестоУстановки = GetCellValueAsString(rowElement, scheme, "Место установки");

            return row;
        }

        private string GetCellValueAsString(RowElement row, Scheme scheme, string fieldName)
        {
            var param = scheme.Parameters.FirstOrDefault(p => p.Name == fieldName);
            if (param == null) return null;
            var cell = row.GetCell(param);
            return cell?.Value?.ToString();
        }

        private double? GetCellValueAsDouble(RowElement row, Scheme scheme, string fieldName)
        {
            var param = scheme.Parameters.FirstOrDefault(p => p.Name == fieldName);
            if (param == null) return null;
            var cell = row.GetCell(param);
            if (cell?.Value == null) return null;
            string strVal = cell.Value.ToString().Replace(',', '.');
            if (double.TryParse(strVal, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return null;
        }

        private int? GetCellValueAsInt(RowElement row, Scheme scheme, string fieldName)
        {
            var param = scheme.Parameters.FirstOrDefault(p => p.Name == fieldName);
            if (param == null) return null;
            var cell = row.GetCell(param);
            if (cell?.Value == null || cell.Value is DBNull) return null;
            if (int.TryParse(cell.Value.ToString(), out int result))
            {
                return result;
            }
            return null;
        }

        private string GetTextVariableValue(Document doc, string variableName)
        {
            Variable var = doc.FindVariable(variableName);
            return var?.TextValue;
        }

        private int? GetIntVariableValue(Document doc, string variableName)
        {
            Variable var = doc.FindVariable(variableName);
            if (var == null) return null;
            return (int?)var.RealValue;
        }
    }
}
