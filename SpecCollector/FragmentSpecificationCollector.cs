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

        private static readonly HashSet<string> MullionFragmentNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Стойка.grb", "Стойка2.grb", "Стойка2 левая.grb",
            "Стойка2 правая передпоследняя.grb", "Стойка2 правая крайняя.grb"
        };

        private string _rootПлоскость;
        private int? _rootЭтаж;
        private int? _rootЭтажей;

        private FloorMatcher _floorMatcher;
        private PhaseReader _phaseReader;
        private Dictionary<string, int> _currentDynamicColumns;

        private string _path;
        private string _specCollectorPath;
        private string _specCollectorStagesPath;
        private string _allSpecsPath;
        private LogService _logService;
        private readonly List<FileStream> _fileLocks = new List<FileStream>();

        public FragmentSpecificationCollector()
        {
            Document doc = TFlex.Application.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("No active T-FLEX document.");
            }

            _path = doc.FilePath;

            string directory = Path.GetDirectoryName(_path);
            if (string.IsNullOrEmpty(directory))
                throw new ArgumentException("Не удалось определить директорию для пути: " + _path);

            _specCollectorPath = Path.Combine(directory, SpecData.SpecCollectorFileName);
            _specCollectorStagesPath = Path.Combine(directory, SpecData.SpecCollectorStagesFileName);
            _allSpecsPath = Path.Combine(directory, SpecData.AllSpecsFileName);

            ___DisplayService.Show();
            ___DisplayService.Log("Проверка входных файлов...");

            if (!ExcelExporter.ValidateInputFiles(directory))
                throw new InvalidOperationException("Входные файлы недоступны");
        }

        /// <summary>
        /// Блокирует выходные файлы монопольно до конца работы.
        /// </summary>
        private void LockOutputFiles(IEnumerable<string> filePaths)
        {
            foreach (var path in filePaths)
            {
                var stream = ExcelExporter.LockFile(path);
                if (stream != null)
                {
                    _fileLocks.Add(stream);
                }
            }
        }

        /// <summary>
        /// Снимает блокировки с выходных файлов.
        /// </summary>
        public void ReleaseLocks()
        {
            foreach (var stream in _fileLocks)
            {
                try { stream.Close(); } catch { }
                try { stream.Dispose(); } catch { }
            }
            _fileLocks.Clear();
        }

        public void GenerateSpec()
        {
            Document doc = TFlex.Application.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("No active T-FLEX document.");
            }

            ___DisplayService.Show();
            ___DisplayService.Log("Старт GenerateSpec");

            _path = doc.FilePath;

            // Проверка выходных файлов
            ___DisplayService.Log("Проверка выходных файлов...");
            if (!ExcelExporter.ValidateFileCanBeWritten(_allSpecsPath, SpecData.AllSpecsFileName) ||
                !ExcelExporter.ValidateFileCanBeWritten(_specCollectorStagesPath, SpecData.SpecCollectorStagesFileName))
            {
                throw new InvalidOperationException("Выходные файлы недоступны для записи");
            }

            _logService = new LogService(_path);

            // Инициализируем FloorMatcher и PhaseReader
            _floorMatcher = new FloorMatcher(_path);
            _phaseReader = new PhaseReader(_path);

            string fn_AllSpecs = _allSpecsPath;

            var allData = new List<Dictionary<string, object>>();

            CollectSpecData(SpecData.Изделия1, allData);
            CollectSpecData(SpecData.Изделия2, allData);
            CollectSpecData(SpecData.Изделия3, allData);
            CollectSpecData(SpecData.Изделия4, allData);

            ExportToMultipleSheets(fn_AllSpecs, allData);

            LockOutputFiles(new[] { _allSpecsPath });

            ExportStagesSpec();

            LockOutputFiles(new[] { _specCollectorStagesPath });

            ReleaseLocks();

            ___DisplayService.Log($"Завершено. Строк: {allData.Count}");
            ___DisplayService.Log("--- LogService ---");
            ___DisplayService.Log(_logService.GetLogText());

            _logService.Flush();
        }

        private void CollectSpecData(SpecData.SpecItem изделие, List<Dictionary<string, object>> allData)
        {
            string dpath = Path.Combine(_path, изделие.FileName);

            Document doc = TFlex.Application.OpenDocument(dpath);
            if (doc == null)
            {
                return;
            }

            var uniqueFloors = _floorMatcher.GetUniqueFloorList(изделие.Плоскость);
            var floorStats = _floorMatcher.GetTypicalFloorStats(изделие.Плоскость, SpecData.FloorRangeFrom, SpecData.FloorRangeTo);

            foreach (var floor in uniqueFloors)
            {
                int count = floorStats.FirstOrDefault(kvp => kvp.Key == floor).Value;

                ___DisplayService.Log($"Set variables: плоскость={изделие.Плоскость}, этаж={floor}, этажей={count}");
                doc.BeginChanges("");
                SetTextVariableValue(doc, SpecData.плоскостьVarName, изделие.Плоскость);
                SetIntegerVariableValue(doc, SpecData.ЭтажVarName, floor);
                SetIntegerVariableValue(doc, SpecData.ЭтажейVarName, count);

                doc.EndChanges();
                doc.Changed = true;

                doc.BeginChanges("");
                UpdateProductStructure(doc);
                doc.EndChanges();

                ___DisplayService.Log($"Формирование спецификации: плоскость={изделие.Плоскость}, этаж={floor}, этажей={count}");

                // Вычисляем динамические столбцы для текущего этажа
                var dynamicColumns = ComputeDynamicColumns(изделие.Плоскость, floor);

                CollectFromDocument(doc, изделие.Плоскость, floor, count, dynamicColumns);

                foreach (var row in _results)
                {
                    allData.Add(row.ToExcelRow());
                }
                _results.Clear();
            }

            doc.Close();
        }

        private Dictionary<string, int> ComputeDynamicColumns(string plane, int floor)
        {
            var result = new Dictionary<string, int>();

            if (_floorMatcher == null || string.IsNullOrEmpty(plane))
                return result;

            // Получаем типовой этаж для текущего этажа и плоскости
            var typicalFloor = _floorMatcher.GetTypicalFloor(plane, floor);
            if (!typicalFloor.HasValue)
                return result;

            // Для каждого RequestItem считаем количество вхождений типового этажа в диапазоне
            if (SpecData.RequestItems != null)
            {
                foreach (var item in SpecData.RequestItems)
                {
                    if (!string.IsNullOrEmpty(item.Name))
                    {
                        int count = _floorMatcher.CountTypicalFloorInRange(plane, item.StartFloor, item.EndFloor, typicalFloor.Value);
                        result[item.Name] = count;
                    }
                }
            }

            return result;
        }

        private void CollectFromDocument(Document doc, string плоскость, int этаж, int этажей, Dictionary<string, int> dynamicColumns = null)
        {
            _rootПлоскость = плоскость;
            _rootЭтаж = этаж;
            _rootЭтажей = этажей;
            _currentDynamicColumns = dynamicColumns ?? new Dictionary<string, int>();
            ProcessDocument(doc, 0, null);
        }

        private void ExportToMultipleSheets(string filePath, List<Dictionary<string, object>> allData)
        {
            var sheets = ExcelExporter.GroupDataBySections(allData);
            var config = new OpenXmlConfiguration()
            {
                FastMode = true,
                EnableAutoWidth = true
            };
            MiniExcel.SaveAs(filePath, sheets, overwriteFile: true, configuration: config);
        }

        private void ExportStagesSpec()
        {
            ___DisplayService.Log("Старт ExportStagesSpec");

            string fn_Stages = _specCollectorStagesPath;
            var stagesData = new List<Dictionary<string, object>>();

            var productItems = new[] { SpecData.Изделия1, SpecData.Изделия2, SpecData.Изделия3, SpecData.Изделия4 };

            foreach (var изделие in productItems)
            {
                string dpath = Path.Combine(_path, изделие.FileName);
                Document doc = TFlex.Application.OpenDocument(dpath);
                if (doc == null)
                {
                    ___DisplayService.Log($"ExportStagesSpec: не удалось открыть {изделие.FileName}");
                    continue;
                }

                var uniqueFloors = _floorMatcher.GetUniqueFloorList(изделие.Плоскость);

            foreach (var typicalFloor in uniqueFloors)
            {
                ___DisplayService.Log($"ExportStagesSpec: плоскость={изделие.Плоскость}, типовой этаж={typicalFloor}");

                doc.BeginChanges("");
                SetTextVariableValue(doc, SpecData.плоскостьVarName, изделие.Плоскость);
                SetIntegerVariableValue(doc, SpecData.ЭтажVarName, typicalFloor);
                SetIntegerVariableValue(doc, SpecData.ЭтажейVarName, 1);
                doc.EndChanges();
                doc.Changed = true;

                doc.BeginChanges("");
                UpdateProductStructure(doc);
                doc.EndChanges();

                // Сбор без динамических столбцов
                CollectFromDocument(doc, изделие.Плоскость, typicalFloor, 1, null);

                    // Определяем физические этажи, соответствующие этому типовому
                    var physicalFloors = Enumerable.Range(SpecData.FloorRangeFrom, SpecData.FloorRangeTo - SpecData.FloorRangeFrom + 1)
                        .Where(f => _floorMatcher.GetTypicalFloor(изделие.Плоскость, f) == typicalFloor)
                        .ToList();

                    ___DisplayService.Log($"ExportStagesSpec: физических этажей={physicalFloors.Count}");

                    foreach (var physFloor in physicalFloors)
                    {
                        foreach (var row in _results)
                        {
                            var excelRow = row.ToExcelRowStages();
                            excelRow["Этаж"] = physFloor;

                            // Вычисляем Этап через PhaseReader
                            if (row.MullionNumber.HasValue)
                            {
                                string mullion = "с" + row.MullionNumber.Value.ToString();
                                string phase = _phaseReader.GetPhase(изделие.Плоскость, mullion, physFloor);
                                excelRow["Этап"] = (object)phase ?? DBNull.Value;
                            }

                            stagesData.Add(excelRow);
                        }
                    }

                    _results.Clear();
                }

                doc.Close();
            }

            ___DisplayService.Log($"ExportStagesSpec: всего строк={stagesData.Count}");

            if (stagesData.Count > 0)
            {
                var sheets = ExcelExporter.GroupDataBySections(stagesData);
                var config = new OpenXmlConfiguration()
                {
                    FastMode = true,
                    EnableAutoWidth = true
                };
                MiniExcel.SaveAs(fn_Stages, sheets, overwriteFile: true, configuration: config);
                ___DisplayService.Log($"ExportStagesSpec: файл сохранён {fn_Stages}");
            }
            else
            {
                ___DisplayService.Log("ExportStagesSpec: нет данных для экспорта");
            }
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

            CollectAndExport(doc);
        }

        public void CollectAndExport(Document rootDocument)
        {
            ___DisplayService.Show();
            ___DisplayService.Log("Старт CollectAndExport");

            // Проверка выходного файла
            if (!ExcelExporter.ValidateFileCanBeWritten(_specCollectorPath, SpecData.SpecCollectorFileName))
                throw new InvalidOperationException("Выходной файл недоступен для записи");

            _results.Clear();
            _logService = new LogService(rootDocument.FilePath);

            // Инициализируем FloorMatcher для работы с Этажи.xlsx
            _floorMatcher = new FloorMatcher(rootDocument.FilePath);

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

                    ___DisplayService.Log($"Сбор: плоскость={targetItem.Плоскость}, этаж={floor.Этаж}, этажей={floor.Этажей}");

                    var dynamicColumns = ComputeDynamicColumns(targetItem.Плоскость, floor.Этаж);
                    CollectFromDocument(rootDocument, targetItem.Плоскость, floor.Этаж, floor.Этажей, dynamicColumns);
                }
            }
            else
            {
                _rootПлоскость = GetTextVariableValue(rootDocument, SpecData.плоскостьVarName);
                _rootЭтаж = GetIntVariableValue(rootDocument, SpecData.ЭтажVarName);

                ___DisplayService.Log($"Сбор (fallback): плоскость={_rootПлоскость}, этаж={_rootЭтаж}");

                var dynamicColumns = ComputeDynamicColumns(_rootПлоскость ?? "", _rootЭтаж ?? 0);
                _currentDynamicColumns = dynamicColumns ?? new Dictionary<string, int>();

                ProcessDocument(rootDocument, 0);
            }

            var exporter = new ExcelExporter(_specCollectorPath);
            exporter.Export(_results);

            LockOutputFiles(new[] { _specCollectorPath });

            // Этапы — раскрытие из уже собранных _results
            _phaseReader = new PhaseReader(_path);
            var stagesData = new List<Dictionary<string, object>>();
            foreach (var row in _results)
            {
                if (!row.Этаж.HasValue) continue;
                var physFloors = Enumerable.Range(SpecData.FloorRangeFrom, SpecData.FloorRangeTo - SpecData.FloorRangeFrom + 1)
                    .Where(f => _floorMatcher.GetTypicalFloor(row.Плоскость, f) == row.Этаж.Value)
                    .ToList();
                foreach (var physFloor in physFloors)
                {
                    var excelRow = row.ToExcelRowStages();
                    excelRow["Этаж"] = physFloor;
                    if (row.MullionNumber.HasValue)
                    {
                        string phase = _phaseReader.GetPhase(row.Плоскость, "с" + row.MullionNumber.Value, physFloor);
                        excelRow["Этап"] = (object)phase ?? DBNull.Value;
                    }
                    stagesData.Add(excelRow);
                }
            }
            ___DisplayService.Log($"Строк этапов: {stagesData.Count}");
            if (stagesData.Count > 0)
            {
                var sheets = ExcelExporter.GroupDataBySections(stagesData);
                var cfg = new OpenXmlConfiguration() { FastMode = true, EnableAutoWidth = true };
                MiniExcel.SaveAs(_specCollectorStagesPath, sheets, overwriteFile: true, configuration: cfg);
            }

            LockOutputFiles(new[] { _specCollectorStagesPath });

            ReleaseLocks();

            ___DisplayService.Log($"Завершено.");
            ___DisplayService.Log("--- LogService ---");
            ___DisplayService.Log(_logService.GetLogText());

            _logService.Flush();
        }

        private void ProcessDocument(Document doc, int level, List<string> parentChain = null, bool isRigelBranch = false, int? mullionNumber = null)
        {
            if (doc == null) return;

            if (parentChain == null) parentChain = new List<string>();

            // Получить отображаемое имя текущего документа и добавить в цепочку
            string currentDisplayName = GetDocumentDisplayName(doc);
            parentChain.Add(currentDisplayName);

            // Если текущий фрагмент — стойка, читаем переменную "Стойка"
            string docFileName = Path.GetFileName(doc.FileName ?? "");
            if (!mullionNumber.HasValue && MullionFragmentNames.Contains(docFileName))
            {
                try
                {
                    var rackVar = doc.FindVariable("Стойка");
                    if (rackVar != null)
                    {
                        mullionNumber = (int)rackVar.RealValue;
                    }
                }
                catch { /* переменная Стойка может отсутствовать */ }
            }

            // Обновляем флаг: если текущий файл — Ригель.grb, ригель2.grb или Ригель гнутый левый.grb, включаем флаг для всей ветви
            if (!isRigelBranch &&
                (docFileName.IndexOf("Ригель", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 docFileName.IndexOf("ригель2", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                isRigelBranch = true;
            }

            var productStructures = doc.GetProductStructures();
            if (productStructures == null || productStructures.Count == 0)
            {
                return;
            }

            var productStructure = productStructures.FirstOrDefault(
                t => t?.GetName(ModelObjectName.Name) == ProductStructureName);
            if (productStructure == null)
            {
                return;
            }

            var scheme = productStructure.GetScheme();
            var allRows = productStructure.GetAllRowElements();
            if (allRows == null)
            {
                return;
            }

            // Коллекция одобренных фрагментов (прошли IncludeInDoc) для фильтрации в ЭТАПЕ 2
            var approvedFragments = new HashSet<object>(new ReferenceComparer());

            // ЭТАП 1: Проход по BOM-строкам (один проход)
            foreach (var row in allRows)
            {
                // Фильтрация ветвей: если строка не включена в спецификацию — пропускаем
                if (!IsInSpecification(row))
                {
                    continue;
                }

                // Если строка прошла фильтрацию и имеет ссылку на фрагмент — добавляем в одобренные
                if (row.SourceFragmentFirstLevel != null)
                {
                    approvedFragments.Add(row.SourceFragmentFirstLevel);
                }

                // Определяем происхождение записи
                if (row.SourceRowElementUID == Guid.Empty)
                {
                    // Родная строка → добавляем в результат
                    var specRow = ExtractRowFromRowElement(row, scheme, doc, isRigelBranch, mullionNumber);
                    if (specRow != null)
                    {
                        _results.Add(specRow);
                    }
                }
                // Заимствованные строки пропускаются (обработка через ЭТАП 2)
            }

            // ЭТАП 2: Проход по фрагментам документа (один проход)
            // Фильтрация: обходим только те фрагменты, чьи BOM-строки прошли IncludeInDoc в ЭТАПЕ 1
            var fragments = doc.GetFragments();
            if (fragments != null)
            {
                foreach (var frag in fragments)
                {
                    if (string.IsNullOrEmpty(frag.FilePath))
                    {
                        continue;
                    }

                    // Проверка: фрагмент должен быть в одобренных
                    if (!approvedFragments.Contains(frag))
                    {
                        continue;
                    }

                    try
                    {
                        frag.Regenerate(true);
                        Document subDoc = frag.GetFragmentDocument(true);

                        if (subDoc != null)
                        {
                            ProcessDocument(subDoc, level + 1, new List<string>(parentChain), isRigelBranch, mullionNumber);
                        }
                    }
                    catch { /* фрагмент может быть недоступен для регенерации */ }
                }
            }

            // Удалить текущий объект из цепочки перед выходом
            if (parentChain.Count > 0)
                parentChain.RemoveAt(parentChain.Count - 1);
        }

        private bool IsInSpecification(RowElement rowElement)
        {
            var incDoc = rowElement.IncludeInDoc?.Value;
            return incDoc != null && (bool)incDoc;
        }

        private SpecificationRow ExtractRowFromRowElement(RowElement rowElement, Scheme scheme, Document doc, bool isRigelBranch, int? mullionNumber = null)
        {
            var row = new SpecificationRow
            {
                Плоскость = _rootПлоскость,
                Этаж = _rootЭтаж,
                Этажей = _rootЭтажей,
                Источник = Path.GetFileName(doc.FileName ?? ""),
                DynamicColumns = new Dictionary<string, int>(_currentDynamicColumns),
                MullionNumber = mullionNumber
            };

            row.Артикул = GetCellValueAsString(rowElement, scheme, "Артикул");
            row.АртикулБазовый = GetCellValueAsString(rowElement, scheme, "Артикул базовый");
            row.КодЗаказа = GetCellValueAsString(rowElement, scheme, "Код заказа");
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
            row.МестоУстановки = GetCellValueAsString(rowElement, scheme, "Место установки");

            // Размещение: Ригель, если хоть один предок в ветви содержит "Ригель" или "Ригель2"
            row.Размещение = isRigelBranch ? "Ригель" : "Стойка";

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
            catch { /* expected: переменная $Наименование может отсутствовать */ }
            return Path.GetFileName(doc.FileName ?? "Unknown");
        }

        private int? GetIntVariableValue(Document doc, string variableName)
        {
            Variable var = doc.FindVariable(variableName);
            if (var == null) return null;
            return (int?)var.RealValue;
        }

        //public List<SpecificationRow> Collect(Document rootDocument)
        //{
        //    ___DisplayService.Show();
        //    ___DisplayService.Log("Старт Collect");
        //
        //    _results.Clear();
        //    _logService = new LogService(rootDocument.FilePath);
        //
        //    // Находим подходящий SpecItem по имени файла документа
        //    string docFileName = Path.GetFileName(rootDocument.FilePath);
        //    var specItems = new[] { SpecData.Изделия1, SpecData.Изделия2, SpecData.Изделия3, SpecData.Изделия4 };
        //    SpecData.SpecItem targetItem = default;
        //    bool found = false;
        //    foreach (var item in specItems)
        //    {
        //        if (string.Equals(item.FileName, docFileName, StringComparison.OrdinalIgnoreCase))
        //        {
        //            targetItem = item;
        //            found = true;
        //            break;
        //        }
        //    }
        //
        //    if (found && targetItem.Этажи != null && targetItem.Этажи.Length > 0)
        //    {
        //        foreach (var floor in targetItem.Этажи)
        //        {
        //            rootDocument.BeginChanges("");
        //            SetTextVariableValue(rootDocument, SpecData.плоскостьVarName, targetItem.Плоскость);
        //            SetIntegerVariableValue(rootDocument, SpecData.ЭтажVarName, floor.Этаж);
        //            SetIntegerVariableValue(rootDocument, SpecData.ЭтажейVarName, floor.Этажей);
        //            rootDocument.EndChanges();
        //            rootDocument.Changed = true;
        //
        //            rootDocument.BeginChanges("");
        //            UpdateProductStructure(rootDocument);
        //            rootDocument.EndChanges();
        //
        //            ___DisplayService.Log($"Сбор: плоскость={targetItem.Плоскость}, этаж={floor.Этаж}, этажей={floor.Этажей}");
        //
        //            // Вычисляем динамические столбцы для текущего этажа
        //            var dynamicColumns = ComputeDynamicColumns(targetItem.Плоскость, floor.Этаж);
        //
        //            CollectFromDocument(rootDocument, targetItem.Плоскость, floor.Этаж, floor.Этажей, dynamicColumns);
        //        }
        //    }
        //    else
        //    {
        //        _rootПлоскость = GetTextVariableValue(rootDocument, SpecData.плоскостьVarName);
        //        _rootЭтаж = GetIntVariableValue(rootDocument, SpecData.ЭтажVarName);
        //        _rootЭтажей = GetIntVariableValue(rootDocument, SpecData.ЭтажейVarName);
        //
        //        ___DisplayService.Log($"Сбор (fallback): плоскость={_rootПлоскость}, этаж={_rootЭтаж}");
        //
        //        // Для fallback тоже вычисляем динамические столбцы
        //        var dynamicColumns = ComputeDynamicColumns(_rootПлоскость ?? "", _rootЭтаж ?? 0);
        //        _currentDynamicColumns = dynamicColumns ?? new Dictionary<string, int>();
        //
        //        ProcessDocument(rootDocument, 0);
        //    }
        //
        //    ___DisplayService.Log($"Завершено. Строк: {_results.Count}");
        //    ___DisplayService.Log("--- LogService ---");
        //    ___DisplayService.Log(_logService.GetLogText());
        //
        //    _logService.Flush();
        //    return new List<SpecificationRow>(_results);
        //}

    }
}