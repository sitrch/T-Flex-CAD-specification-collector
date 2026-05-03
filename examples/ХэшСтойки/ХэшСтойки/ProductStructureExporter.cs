using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using TFlex;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;
using TFlex.Model.Model3D;
using static TFlex.Model.RowElementGroup;
using static TFlex.Model.Units.StandardUnits;
using System.Collections.Generic;




namespace TFAPIHash
{
    public class ProductStructureExporter
    {
        public Document document;
        public string path;

        public ProductStructureExporter()
        {
            this.document = TFlex.Application.ActiveDocument;
            this.path = Path.Combine(document.FilePath, "ProductStructure.xlsx");
        }
        public ProductStructureExporter(Document document, string path)
        {
            this.document = document;
            this.path = path;
        }

        public void export()
        {
            ExcelMiniDataBlock block = new ExcelMiniDataBlock(path, "");
            block.CreateEmpty();

            //ExportProductStructure(Data.Раскрой,dataBlock);

            MessageBox.Show("stub");
        }
        public void ExportProductStructure(Data.view ViewData, ref ExcelMiniDataBlock dataBlock, Dictionary<string, object> ColumnsPlusDict)
        {
            ExcelExportConfiguration exco = new ExcelExportConfiguration(document, ViewData.СоставИзделия, ViewData.Представление, ViewData.columns);
            //List<ExcelMiniDataBlock> dataBlockList = new List<ExcelMiniDataBlock>(); 
            exco.UpdateStructure();

            if (dataBlock == null)
            {
                dataBlock = new ExcelMiniDataBlock(path, ViewData.Представление);
            }


            // Здесь обработка каждой группы строк (верхнего уровня) педставления
            //ExcelMiniDataBlock block = null;

            //if (!ViewData.OneGroupOneSheet) 
            //{

            // Все группы на одном листе
            //dataBlockList.Add(block);
            //}
            // Каждая группа на отдельном листе
            foreach (RowElementGroup rgroup in exco.rgroups)
            {
                //if (ViewData.OneGroupOneSheet) 
                //{
                //block = new ExcelMiniDataBlock(path, rgroup.Name);
                //dataBlockList.Add(block);
                //}
                ProcessGroup(exco, rgroup, dataBlock, ColumnsPlusDict);
            }

        }  
        
        // Получить каждую строку из группы представления
        void ProcessGroup(ExcelExportConfiguration exco, RowElementGroup group, ExcelMiniDataBlock dblock,Dictionary<string, object> AddColumnsDict)
        {
            foreach (RowElementGroup.Item item in group.Items)
            {
                Dictionary<Guid, RowElementGroup.MergedCellValue> MergedCells = item.MergedCells;
                List<RowElement> MergedElements = item.MergedElements;

                var excelRow = new Dictionary<string, object>();

                // Добавляем данные вручную
                foreach (KeyValuePair<string, object> kvp in AddColumnsDict)
                {
                    if (excelRow.ContainsKey(kvp.Key))
                    {
                        throw new ArgumentException("Ключ уже существует: " + kvp.Key);
                    }

                    excelRow.Add(kvp.Key, kvp.Value);
                }

                //Заполняем данными спецификации
                foreach (ParameterDescriptor pd in exco.pColumns)
                {
                    excelRow[pd.Name] = exco.GetCellValue(pd, item);

                }

                excelRow.Add("Общее количество", (int?)null);

                dblock.AddRow(excelRow);
            }
        }
        public void DeleteFile(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (UnauthorizedAccessException)
            {
                // Ошибка: нет прав или файл помечен "Только для чтения"
                Console.WriteLine("Доступ запрещен. Запустите от имени администратора.");
            }
            catch (IOException)
            {
                // Ошибка: файл открыт в другой программе
                Console.WriteLine("Файл занят другим процессом.");
            }
            catch (Exception ex)
            {
                // Все остальные непредвиденные ошибки
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }
    }
}
