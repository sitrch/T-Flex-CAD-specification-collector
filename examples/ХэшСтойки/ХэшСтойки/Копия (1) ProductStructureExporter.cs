using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using TFlex;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;
using static TFlex.Model.RowElementGroup;
using static TFlex.Model.Units.StandardUnits;
using MiniExcelLibs.Attributes;




namespace TFAPIHash
{
    public class ProductStructureExporter
    {
        static Document document = TFlex.Application.ActiveDocument;
        string path = Path.Combine(document.FilePath, "ProductStructure.xlsx");
        ProductStructure ps;
        List<ParameterDescriptor> pColumns = new List<ParameterDescriptor>();

        public void export()
        {
            ExportProductStructure(document);
        }
        public void ExportProductStructure(Document doc)
        {
            var dataList = new List<Dictionary<string, object>>();

            var ps = document.GetProductStructures()
                 .FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == "m2spec");
            if (ps == null) return;

            document.BeginChanges("");
            ps.UpdateStructure();
            document.EndChanges();

            Scheme scheme = ps.GetScheme();

            AddParameter("Артикул", scheme, pColumns);
            AddParameter("Наименование", scheme, pColumns);
            AddParameter("Длина", scheme, pColumns);
            AddParameter("Количество всего", scheme, pColumns);

            Groupings groupings = scheme.Properties.Groupings;
            GroupingRules myRules = groupings.FirstOrDefault(g => g.Name == "Спецификация");

            if (myRules == null)
            {
                throw new Exception("Grouping rule not found");    
            }

            Guid id = myRules.ID;
            // Получаем все элементы строк
            ICollection<RowElementGroup> rgroups = ps.GetRowElementGroups(id);

            // Здесь обработка каждой группы строк (верхнего уровня) педставления
            foreach (RowElementGroup rgroup in rgroups)
            {
                ProcessGroup(rgroup, dataList);
            }

            var config = new OpenXmlConfiguration()
            {
                FastMode = true, // 1. Обязательно включаем FastMode для работы автоширины
                EnableAutoWidth = true // Включает автоподбор ширины по содержимому
            };
            // 4. Сохраняем результат в Excel (MiniExcel сделает всё сам)
            MiniExcel.SaveAs(path, dataList, sheetName: "Состав", overwriteFile: true, configuration: config);
        }
        
        // Получить каждую строку из группы представления
        void ProcessGroup(RowElementGroup group, List<Dictionary<string, object>> resultList)
        {
            foreach (RowElementGroup.Item item in group.Items)
            {
                Dictionary<Guid, RowElementGroup.MergedCellValue> MergedCells = item.MergedCells;
                List<RowElement> MergedElements = item.MergedElements;

                var excelRow = new Dictionary<string, object>();

                foreach (ParameterDescriptor pd in pColumns)
                {
                    object rawValue = null;
                    BomValueType ValueType = pd.ValueType;
                    if (MergedCells.TryGetValue(pd.UID, out var mergedValue))
                    {
                        // Значение общее для всей группы (объединено)
                        rawValue = mergedValue.Value;
                    }
                    else
                    {
                        // Значение не объединено. Извлекаем его из первого элемента группы
                        // MergedElements — это список строк RowElement, входящих в этот item
                        var firstElement = MergedElements.FirstOrDefault();
                        if (firstElement != null)
                        {
                            // Используем дескриптор для получения точного значения ячейки
                            rawValue = firstElement.GetCell(pd)?.Value;
                        }
                    }
                    excelRow[pd.Name] = PrepareValueByDescriptor(rawValue, pd);
                }
                resultList.Add(excelRow);
            }
        }

        private object PrepareValueByDescriptor(object rawValue, ParameterDescriptor pd)
        {
            if (rawValue == null || rawValue is DBNull) return null;

            // В T-FLEX API тип параметра хранится в pd.ValueType (тип BomValueType)
            switch (pd.ValueType)
            {
                case BomValueType.vtInt:
                    if (int.TryParse(rawValue.ToString(), out int intVal))
                        return intVal;
                    return 0; // Или null, если данные некорректны

                case BomValueType.vtReal:
                    // Универсальный парсинг числа (с учетом точки/запятой)
                    string strReal = rawValue.ToString().Replace(',', '.');
                    if (double.TryParse(strReal, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double doubleVal))
                        return doubleVal;
                    return 0.0;

                case BomValueType.vtBool:
                    if (bool.TryParse(rawValue.ToString(), out bool boolVal))
                        return boolVal;
                    // T-FLEX иногда отдает "0"/"1" для Bool
                    return rawValue.ToString() == "1";

                case BomValueType.vtString:
                    return rawValue.ToString();

                case BomValueType.vtGuid:
                    return rawValue is Guid ? rawValue : rawValue.ToString();

                case BomValueType.vtAuto:
                    // Если тип авто, пробуем угадать или отдаем как есть
                    return rawValue;

                case BomValueType.vtBinary:
                    return "[Binary Data]"; // Бинарные данные в Excel обычно не пишут

                default:
                    return rawValue.ToString();
            }
        }
        public void AddParameter(string name, Scheme scheme, List<ParameterDescriptor> List)
        {
            ParameterDescriptor parameter = scheme.Parameters.FirstOrDefault(p => p.Name == name);
            if (parameter != null)
            {
                List.Add(parameter);
            }
        }
    }
}
