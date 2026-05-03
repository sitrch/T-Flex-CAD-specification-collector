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
        
        private string _rootПлоскость;
        private int? _rootЭтаж;
        private int? _rootЭтажей;

        public bool OnlyBOMGenerate = false;
        public bool IncludeAllRows = false; // Если true — отключает фильтрацию по IncludeInDoc (для отладки)
        public string LogFilterArticul = null; // Если задано, логируются только строки с этим артикулом
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

            System.IO.File.WriteAllText("collector.log1", $"[{DateTime.Now}] ========== START GenerateSpec ==========\n");
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] IncludeAllRows={IncludeAllRows}\n");
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Folder: {_path}\n");

            string fn_AllSpecs = Path.Combine(_path, "Спецификации.xlsx");

            var allData = new List<Dictionary<string, object>>();

            CollectSpecData(SpecData.Изделия1, allData);
            CollectSpecData(SpecData.Изделия2, allData);
            CollectSpecData(SpecData.Изделия3, allData);
            CollectSpecData(SpecData.Изделия4, allData);

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Total _results count before export: {_results.Count}\n");
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Total allData count: {allData.Count}\n");

            ExportToMultipleSheets(fn_AllSpecs, allData);
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Done! File: {fn_AllSpecs}\n");
        }

        private void CollectSpecData(SpecData.SpecItem изделие, List<Dictionary<string, object>> allData)
        {
            string dpath = Path.Combine(_path, изделие.Плоскость) == null ? изделие.FileName : Path.Combine(_path, изделие.FileName);
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] ===== Processing: {изделие.FileName} ({изделие.Плоскость}) =====\n");

            Document doc = TFlex.Application.OpenDocument(dpath);
            if (doc == null)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]  Ошибка: документ не найден\n");
                return;
            }

            foreach (var floor in изделие.Этажи)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] --- {изделие.Плоскость} Floor {floor.Этаж} (of {floor.Этажей}) ---\n");
                doc.BeginChanges("");
                SetTextVariableValue(doc, SpecData.плоскостьVarName, изделие.Плоскость);
                SetIntegerVariableValue(doc, SpecData.ЭтажVarName, floor.Этаж);

                doc.EndChanges();
                doc.Changed = true;

                doc.BeginChanges("");
                UpdateProductStructure(doc);
                doc.EndChanges();

                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Сбор данных ... (current _results={_results.Count})\n");
                CollectFromDocument(doc, изделие.Плоскость, floor.Этаж, floor.Этажей);

                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Adding {_results.Count} rows to allData...\n");
                foreach (var row in _results)
                {
                    allData.Add(row.ToExcelRow());
                    System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    Added: {row.Артикул} | {row.Наименование} | Спецификация={row.Спецификация} | Source={row.Источник}\n");
                }
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] allData total after adding: {allData.Count}\n");
                _results.Clear();
            }

            doc.Close();
        }

        private void CollectFromDocument(Document doc, string плоскость, int этаж, int этажей)
        {
            _rootПлоскость = плоскость;
            _rootЭтаж = этаж;
            _rootЭтажей = этажей;
            ProcessDocument(doc, 0, null);
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

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] === Started {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n");

            if (found && targetItem.Этажи != null && targetItem.Этажи.Length > 0)
            {
                // Перебираем все этажи из SpecItem, как в GenerateSpec()
                foreach (var floor in targetItem.Этажи)
                {
                    System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]  Processing floor: Этаж={floor.Этаж}, Этажей={floor.Этажей}\n");
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

                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Плоскость={_rootПлоскость}, Этаж={_rootЭтаж}\n");
                ProcessDocument(rootDocument, 0);
            }

            var exporter = new ExcelExporter(outputExcelPath);
            exporter.Export(_results);
            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] === Finished: {_results.Count} rows collected ===\n");
        }

        private void ProcessDocument(Document doc, int level, List<string> parentChain = null)
        {
            if (doc == null) return;

            if (parentChain == null) parentChain = new List<string>();

            var docPath = doc.FileName ?? "";
            var docName = Path.GetFileName(docPath);

            // Получить отображаемое имя текущего документа и добавить в цепочку
            string currentDisplayName = GetDocumentDisplayName(doc);
            parentChain.Add(currentDisplayName);

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] >>> ENTER Level={level}, doc={currentDisplayName}, path={docPath}\n");

            var productStructures = doc.GetProductStructures();
            if (productStructures == null || productStructures.Count == 0)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] <<< EXIT Level={level} (no product structures)\n");
                return;
            }

            var productStructure = productStructures.FirstOrDefault(
                t => t?.GetName(ModelObjectName.Name) == ProductStructureName);
            if (productStructure == null)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] <<< EXIT Level={level} (no m2spec)\n");
                return;
            }

            var scheme = productStructure.GetScheme();
            var allRows = productStructure.GetAllRowElements();
            if (allRows == null)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] <<< EXIT Level={level} (no rows)\n");
                return;
            }

            int ownCount = 0, borrowedCount = 0, skippedSpec = 0;

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Level={level}, allRows={allRows.Count}\n");

            // ЭТАП 1: Проход по BOM-строкам (один проход)
            foreach (var row in allRows)
            {
                var rowName = GetCellValueAsString(row, scheme, "Наименование") ?? "(null)";
                var rowArt = GetCellValueAsString(row, scheme, "Артикул") ?? "(null)";

                // Фильтрация ветвей: если строка не включена в спецификацию — пропускаем (кроме режима отладки)
                if (!IncludeAllRows && !IsInSpecification(row))
                {
                    skippedSpec++;
                    if (ShouldLogRow(rowArt))
                    {
                        string chainStr = string.Join(" → ", parentChain);
                        // TFlexAPI monitoring
                        var srcFrag = row.SourceFragmentFirstLevel?.GetType().Name ?? "null";
                        var srcFrag3D = row.SourceFragment3DFirstLevel?.GetType().Name ?? "null";
                        var srcFragPath = row.SourceFragmentPath ?? "null";
                        var srcUIDFirst = row.SourceRowElementUIDFirstLevel.ToString();
                        var srcUID = row.SourceRowElementUID.ToString();
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    [{rowArt}] [{rowName}] SKIPPED (chain={chainStr}, IncludeInDoc=false, srcFrag={srcFrag}, srcFrag3D={srcFrag3D}, srcFragPath={srcFragPath}, srcUIDFirst={srcUIDFirst}, srcUID={srcUID})\n");
                    }
                    continue;
                }


                // Определяем происхождение записи
                if (row.SourceRowElementUID == Guid.Empty)
                {
                    // Родная строка → добавляем в результат
                    if (ShouldLogRow(rowArt))
                    {
                        string chainStr = string.Join(" → ", parentChain);
                        // TFlexAPI monitoring
                        var srcFrag = row.SourceFragmentFirstLevel?.GetType().Name ?? "null";
                        var srcFrag3D = row.SourceFragment3DFirstLevel?.GetType().Name ?? "null";
                        var srcFragPath = row.SourceFragmentPath ?? "null";
                        var srcUIDFirst = row.SourceRowElementUIDFirstLevel.ToString();
                        var srcUID = row.SourceRowElementUID.ToString();
                        var srcObj = row.SourceObject?.GetType().Name ?? "null";
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    [{rowArt}] [{rowName}] OWN ROW (chain={chainStr}, level={level}, srcFrag={srcFrag}, srcFrag3D={srcFrag3D}, srcFragPath={srcFragPath}, srcUIDFirst={srcUIDFirst}, srcUID={srcUID}, srcObj={srcObj})\n");
                    }

                    var specRow = ExtractRowFromRowElement(row, scheme, doc);
                    if (specRow != null)
                    {
                        _results.Add(specRow);
                        ownCount++;
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    --> ADDED to _results (total={_results.Count})\n");
                    }
                }
                else
                {
                    // Заимствованная строка — пропускаем (обработка через ЭТАП 2)
                    borrowedCount++;
                    if (ShouldLogRow(rowArt))
                    {
                        string chainStr = string.Join(" → ", parentChain);
                        // TFlexAPI monitoring
                        var srcFrag = row.SourceFragmentFirstLevel?.GetType().Name ?? "null";
                        var srcFrag3D = row.SourceFragment3DFirstLevel?.GetType().Name ?? "null";
                        var srcFragPath = row.SourceFragmentPath ?? "null";
                        var srcUIDFirst = row.SourceRowElementUIDFirstLevel.ToString();
                        var srcUID = row.SourceRowElementUID.ToString();
                        var srcObj = row.SourceObject?.GetType().Name ?? "null";
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    [{rowArt}] [{rowName}] BORROWED (chain={chainStr}, SRE_UID={row.SourceRowElementUID}, srcFrag={srcFrag}, srcFrag3D={srcFrag3D}, srcFragPath={srcFragPath}, srcUIDFirst={srcUIDFirst}, srcUID={srcUID}, srcObj={srcObj})\n");
                    }
                }
            }

            // ЭТАП 2: Проход по фрагментам документа (один проход)
            var fragments = doc.GetFragments();
            if (fragments != null)
            {
                System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Level={level}: Processing {fragments.Count} fragments...\n");
                foreach (var frag in fragments)
                {
                    if (string.IsNullOrEmpty(frag.FilePath))
                    {
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    Fragment SKIPPED (empty FilePath)\n");
                        continue;
                    }

                    try
                    {
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    Fragment: {frag.FilePath}\n");
                        frag.Regenerate(true);
                        Document subDoc = frag.GetFragmentDocument(true);
                        if (subDoc != null)
                        {
                            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    --> Recursing into: {GetDocumentDisplayName(subDoc)}\n");
                            ProcessDocument(subDoc, level + 1, new List<string>(parentChain));
                        }
                        else
                        {
                            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    --> subDoc is NULL\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}]    --> ERROR: {ex.Message}\n");
                    }
                }
            }

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] Level={level} summary: own={ownCount} borrowed={borrowedCount} skippedSpec={skippedSpec}, total_results={_results.Count}\n");

            // Удалить текущий объект из цепочки перед выходом
            if (parentChain.Count > 0)
                parentChain.RemoveAt(parentChain.Count - 1);

            System.IO.File.AppendAllText("collector.log1", $"[{DateTime.Now}] <<< EXIT Level={level}, doc={currentDisplayName}\n");
        }

        private bool IsInSpecification(RowElement rowElement)
        {
            var incDoc = rowElement.IncludeInDoc?.Value;
            return incDoc != null && (bool)incDoc;
        }

        private SpecificationRow ExtractRowFromRowElement(RowElement rowElement, Scheme scheme, Document doc)
        {
            var name = GetCellValueAsString(rowElement, scheme, "Наименование") ?? "(empty)";

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

        /// <summary>
        /// Получает отображаемое имя документа: сначала пытается получить переменную $Наименование, если нет — возвращает имя файла.
        /// </summary>
        private string GetDocumentDisplayName(Document doc)
        {
            try
            {
                Variable var = doc.FindVariable("$Наименование");
                if (var != null && !string.IsNullOrEmpty(var.TextValue))
                {
                    return var.TextValue;
                }
            }
            catch { }
            return Path.GetFileName(doc.FileName ?? "Unknown");
        }

        /// <summary>
        /// Проверяет, нужно ли логировать строку (по артикулу)
        /// </summary>
        private bool ShouldLogRow(string rowArt)
        {
            if (string.IsNullOrEmpty(LogFilterArticul)) return true; // логировать всё
            if (string.IsNullOrEmpty(rowArt)) return false;
            return rowArt.Contains(LogFilterArticul);
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
