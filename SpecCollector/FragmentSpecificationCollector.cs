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
        private readonly List<SpecificationRow> _results = new List<SpecificationRow>();
        //private readonly HashSet<string> _processedDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private string _rootПлоскость;
        private int? _rootЭтаж;
        private int? _rootЭтажей;

        public bool OnlyBOMGenerate = false;
        private string _path;

        public void GenerateSpec()
        {
            Document doc = TFlex.Application.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("No active T-FLEX document.");
            }

            //OnlyBOMGenerate = true;
            _path = doc.FilePath;

            SpecDebugService.Show();
            SpecDebugService.Log("Start collecting specifications...");
            SpecDebugService.Log("Folder: " + _path);

            string fn_AllSpecs = Path.Combine(_path, "Спецификации.xlsx");

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

            foreach (var floor in изделие.Этажи)
            {
                SpecDebugService.Log($" {изделие.Плоскость} Этаж {floor.Этаж} ...");
                doc.BeginChanges("");
                SetTextVariableValue(doc, SpecData.плоскостьVarName, изделие.Плоскость);
                SetIntegerVariableValue(doc, SpecData.ЭтажVarName, floor.Этаж);

                doc.EndChanges();
                doc.Changed = true;

                doc.BeginChanges("");
                UpdateProductStructure(doc);
                doc.EndChanges();

                SpecDebugService.Log($" Сбор данных ...");
                CollectFromDocument(doc, изделие.Плоскость, floor.Этаж, floor.Этажей);

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
            //_processedDocs.Clear();
            ProcessDocument(doc, 0);
        }

        private void ExportToMultipleSheets(string filePath, List<Dictionary<string, object>> allData)
        {
            var sheets = new Dictionary<string, object>();

            sheets.Add("Спецификация", allData);

            var groupMapping = SpecData.BOMSections;

            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var row in allData)
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

            // Находим подходящий SpecItem по имени файла документа
            string docFileName = Path.GetFileName(rootDocument.FilePath);
            var specItems = new[] { SpecData.Изделия1, SpecData.Изделия2, SpecData.Изделия3, SpecData.Изделия4 };
            SpecData.SpecItem targetItem = default;
            bool found = false;
            foreach (var item in specItems)
            {
                if (string.Equals(item.FileName, docFileName, StringComparison.OrdinalIgnoreCase))
                {
                    targetItem = item;
                    found = true;
                    break;
                }
            }

            SpecDebugService.Show();
            SpecDebugService.Log("=== Started " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ===");

            if (found && targetItem.Этажи != null && targetItem.Этажи.Length > 0)
            {
                // Перебираем все этажи из SpecItem, как в GenerateSpec()
                foreach (var floor in targetItem.Этажи)
                {
                    SpecDebugService.Log("  Processing floor: Этаж=" + floor.Этаж + ", Этажей=" + floor.Этажей);
                    rootDocument.BeginChanges("");
                    SetTextVariableValue(rootDocument, SpecData.плоскостьVarName, targetItem.Плоскость);
                    SetIntegerVariableValue(rootDocument, SpecData.ЭтажVarName, floor.Этаж);
                    SetIntegerVariableValue(rootDocument, SpecData.ЭтажейVarName, floor.Этажей);
                    rootDocument.EndChanges();
                    rootDocument.Changed = true;

                    rootDocument.BeginChanges("");
                    UpdateProductStructure(rootDocument);
                    rootDocument.EndChanges();

                    CollectFromDocument(rootDocument, targetItem.Плоскость, floor.Этаж, floor.Этажей);
                }
            }
            else
            {
                // Fallback: читаем значения из переменных документа
                _rootПлоскость = GetTextVariableValue(rootDocument, SpecData.плоскостьVarName);
                _rootЭтаж = GetIntVariableValue(rootDocument, SpecData.ЭтажVarName);

                SpecDebugService.Log("Плоскость=" + _rootПлоскость + ", Этаж=" + _rootЭтаж);
                ProcessDocument(rootDocument, 0);
            }

            var exporter = new ExcelExporter(outputExcelPath);
            exporter.Export(_results);
            SpecDebugService.Log("=== Finished: " + _results.Count + " rows collected ===");
        }

        private void ProcessDocument(Document doc, int level)
        {
            if (doc == null) return;

            ExtractOwnRecordsFromBOM(doc, level);

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
                    catch (Exception)
                    {
                        // ignore
                    }
                }
            }
        }

        private void ExtractOwnRecordsFromBOM(Document doc, int level)
        {
            var productStructures = doc.GetProductStructures();
            if (productStructures == null || productStructures.Count == 0)
            {
                SpecDebugService.Log("  -> NO ProductStructures found");
                return;
            }


            var productStructure = productStructures?.FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == ProductStructureName);
            if (productStructure == null) return;

            var scheme = productStructure.GetScheme();
            var allRows = productStructure.GetAllRowElements();

            if (allRows != null)
            {
                foreach (var row in allRows)
                {
                    var specRow = ExtractRowFromRowElement(row, scheme, doc);
                    if(specRow != null)
                    {
                        _results.Add(specRow);
                    }

                }
            }
        }

        /// <summary>
        /// Определяет, что запись принадлежит текущему документу (рекомендуемый способ).
        /// Запись считается «своей», если она не поднята из фрагмента (SourceFragmentPath == null)
        /// и при этом собрана по объекту текущего документа (SourceObject != null).
        /// </summary>
        private bool IsOwnBySourcePathAndObject(RowElement rowElement, Document doc)
        {
            return rowElement.SourceFragmentPath == null && rowElement.SourceObject != null;
        }

        /// <summary>
        /// Определяет, что запись принадлежит текущему документу по UID источника.
        /// Если SourceRowElementUID == Guid.Empty, элемент не поднят из фрагмента — своя запись.
        /// Простой способ, не зависит от строковых сравнений путей.
        /// </summary>
        private bool IsOwnBySourceRowElementUID(RowElement rowElement)
        {
            return rowElement.SourceRowElementUID == Guid.Empty;
        }

        /// <summary>
        /// Определяет, что запись принадлежит текущему документу по UID первого уровня вложенности.
        /// Если SourceRowElementUIDFirstLevel == Guid.Empty — элемент не поднят из фрагмента
        /// на первом уровне вложенности. Удобно для многоуровневых сборок, чтобы отличать
        /// прямые дочерние записи от записей, поднятых из глубоко вложенных фрагментов.
        /// </summary>
        private bool IsOwnBySourceFragmentFirstLevel(RowElement rowElement)
        {
            return rowElement.SourceRowElementUIDFirstLevel == Guid.Empty;
        }

        private SpecificationRow ExtractRowFromRowElement(RowElement rowElement, Scheme scheme, Document doc)
        {
            // Диагностическое логирование для анализа свойств RowElement
            var sfp = rowElement.SourceFragmentPath;
            var so = rowElement.SourceObject;
            var sreUid = rowElement.SourceRowElementUID;
            var sreUidFl = rowElement.SourceRowElementUIDFirstLevel;
            var name = GetCellValueAsString(rowElement, scheme, "Наименование") ?? "(empty)";

            if (!IsOwnBySourceRowElementUID(rowElement))
            {
                return null;
            }

            var row = new SpecificationRow
            {
                Плоскость = _rootПлоскость,
                Этаж = _rootЭтаж,
                Этажей = _rootЭтажей,
                Источник = Path.GetFileName(doc.FileName ?? "")
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

        public List<SpecificationRow> Collect(Document rootDocument)
        {
            _results.Clear();

            // Находим подходящий SpecItem по имени файла документа
            string docFileName = Path.GetFileName(rootDocument.FilePath);
            var specItems = new[] { SpecData.Изделия1, SpecData.Изделия2, SpecData.Изделия3, SpecData.Изделия4 };
            SpecData.SpecItem targetItem = default;
            bool found = false;
            foreach (var item in specItems)
            {
                if (string.Equals(item.FileName, docFileName, StringComparison.OrdinalIgnoreCase))
                {
                    targetItem = item;
                    found = true;
                    break;
                }
            }

            if (found && targetItem.Этажи != null && targetItem.Этажи.Length > 0)
            {
                foreach (var floor in targetItem.Этажи)
                {
                    rootDocument.BeginChanges("");
                    SetTextVariableValue(rootDocument, SpecData.плоскостьVarName, targetItem.Плоскость);
                    SetIntegerVariableValue(rootDocument, SpecData.ЭтажVarName, floor.Этаж);
                    SetIntegerVariableValue(rootDocument, SpecData.ЭтажейVarName, floor.Этажей);
                    rootDocument.EndChanges();
                    rootDocument.Changed = true;

                    rootDocument.BeginChanges("");
                    UpdateProductStructure(rootDocument);
                    rootDocument.EndChanges();

                    CollectFromDocument(rootDocument, targetItem.Плоскость, floor.Этаж, floor.Этажей);
                }
            }
            else
            {
                _rootПлоскость = GetTextVariableValue(rootDocument, SpecData.плоскостьVarName);
                _rootЭтаж = GetIntVariableValue(rootDocument, SpecData.ЭтажVarName);
                _rootЭтажей = GetIntVariableValue(rootDocument, SpecData.ЭтажейVarName);
                ProcessDocument(rootDocument, 0);
            }

            return new List<SpecificationRow>(_results);
        }
    }
}
