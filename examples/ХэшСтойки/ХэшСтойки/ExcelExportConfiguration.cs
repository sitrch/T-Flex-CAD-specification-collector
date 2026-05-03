using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;
using static TFlex.Model.RowElementGroup;
using static TFlex.Model.Units.StandardUnits;

namespace TFAPIHash
{
    public class ExcelExportConfiguration
    {
        public Document document;
        public ProductStructure ps;
        public Scheme scheme;
        public GroupingRules groupingRules;
        public List<ParameterDescriptor> pColumns = new List<ParameterDescriptor>();
        public ICollection<RowElementGroup> rgroups;
        public Guid idПравилГруппировки;
        public bool OneGroupOneSheet = false;

        //ExcelMiniDataBlock dataBlock = new ExcelMiniDataBlock(path, Представление);

        public ExcelExportConfiguration(Document document,string СоставИзделия, string Представление, List<string> columns, bool OneGroupOneSheet = false)
        {
            this.document = document;
            this.OneGroupOneSheet = OneGroupOneSheet;

            SetProductStructure(СоставИзделия);
            AddParameters(columns);
            SetGroupingRules(Представление);
            rgroups = ps.GetRowElementGroups(idПравилГруппировки);
        }

        /// <summary>
        /// Столбцы состава изделия, из которых будут собираться данные
        /// </summary>
        /// <param name="columns"></param>
        public void AddParameters(List<string> columns)
        {
            foreach (string column in columns)
            {
                ParameterDescriptor parameter = scheme.Parameters.FirstOrDefault(p => p.Name == column);
                if (parameter == null) 
                { 
                    throw new Exception($"Столбец состава изделия {column} не найден!"); 
                }
                pColumns.Add(parameter);
            }
        }
        public void UpdateStructure()
        {
            document.BeginChanges("");
            ps.MarkChanged();
            ps.UpdateStructure();
            ps.UpdateReports();
            document.EndChanges();

        }
        /// <summary>
        /// Получить значение ячейки из спецификации
        /// </summary>
        /// <param name="pd"> Список столбцов, из которых будут извлекаться данные</param>
        /// <param name="item">Строка с данными спецификации</param>
        /// <returns></returns>
        public object GetCellValue(ParameterDescriptor pd, Item item)
        {
            object rawValue = null;

            Dictionary<Guid, RowElementGroup.MergedCellValue> MergedCells = item.MergedCells;
            List<RowElement> MergedElements = item.MergedElements;

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
            return (PrepareValueByDescriptor(rawValue, pd));
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
        public void SetProductStructure(string СоставИзделия)
        {
            ps = document.GetProductStructures()
                 .FirstOrDefault(t => t?.GetName(ModelObjectName.Name) == СоставИзделия);
            if (ps == null)
            {
                throw new Exception($"Состав изделия {СоставИзделия} не найден!");
            }

            this.scheme = ps.GetScheme();
        }

        public void SetGroupingRules(string Представление)
        {
            groupingRules = scheme.Properties.Groupings.FirstOrDefault(g => g.Name == Представление);

            if (groupingRules == null)
            {
                throw new Exception($"Представление {Представление} не найдено!");
            }
            idПравилГруппировки = groupingRules.ID;
        }

        // Получить каждую строку из группы представления
        void ProcessGroup(RowElementGroup group, ExcelMiniDataBlock dblock)
        {
            foreach (RowElementGroup.Item item in group.Items)
            {
                Dictionary<Guid, RowElementGroup.MergedCellValue> MergedCells = item.MergedCells;
                List<RowElement> MergedElements = item.MergedElements;

                var excelRow = new Dictionary<string, object>();

                foreach (ParameterDescriptor pd in pColumns)
                {
                    excelRow[pd.Name] = GetCellValue(pd, item);

                }
                dblock.AddRow(excelRow);
            }
        }
    }

    public class ExcelMiniDataBlock
    {
        public string path;
        public string SheetName;
        public List<Dictionary<string, object>> DataList;
        OpenXmlConfiguration config;

        public ExcelMiniDataBlock(string path, string SheetName)
        {
            this.path = path;
            this.SheetName = SheetName;

            DataList = new List<Dictionary<string, object>>();

            config = new OpenXmlConfiguration()
            {
                FastMode = true, // 1. Обязательно включаем FastMode для работы автоширины
                EnableAutoWidth = true // Включает автоподбор ширины по содержимому
            };
        }
       
        public void Write(bool OverwriteFile = false)
        {
            // 4. Сохраняем результат в Excel (MiniExcel сделает всё сам)
            MiniExcel.Insert(path, DataList, sheetName: SheetName, configuration: config);
        }

        public void Write(List<ExcelMiniDataBlock> dataBlockList, bool OverwriteFile = false)
        {
            foreach (ExcelMiniDataBlock db in dataBlockList)
            {
                db.Write(OverwriteFile);
            }

        }
            
        public void AddRow(Dictionary<string, object> excelRow)
        {
            DataList.Add(excelRow);
        }

        public void CreateEmpty()
        {
            var columns = new List<Dictionary<string, object>>();
            // В данном случае файл будет действительно пустым, 
            // так как ключей (заголовков) нет в списке.
            MiniExcel.SaveAs(path, columns, overwriteFile: true);
        }
    }
}
