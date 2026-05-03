using MiniExcelLibs;
using MiniExcelLibs.Attributes;
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
using TFlex.Model.Model3D;
using static TFlex.Model.RowElementGroup;
using static TFlex.Model.Units.StandardUnits;




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

            ExportProductStructure(Data.Раскрой);
        }
        public void ExportProductStructure(Data.view ViewData)
        {
            ExcelExportConfiguration exco = new ExcelExportConfiguration(document, ViewData.СоставИзделия, ViewData.Представление, ViewData.columns);
            List<ExcelMiniDataBlock> dataBlockList = new List<ExcelMiniDataBlock>(); 
            exco.UpdateStructure();

            // Здесь обработка каждой группы строк (верхнего уровня) педставления
            ExcelMiniDataBlock block = null;

            if (!ViewData.OneGroupOneSheet) 
            {
                block = new ExcelMiniDataBlock(path, ViewData.Представление);
                dataBlockList.Add(block);
            }
            foreach (RowElementGroup rgroup in exco.rgroups)
            {
                if (ViewData.OneGroupOneSheet) 
                {
                    block = new ExcelMiniDataBlock(path, rgroup.Name);
                    dataBlockList.Add(block);
                }
                ProcessGroup(exco, rgroup, block);
            }

            // 4. Сохраняем результат в Excel (MiniExcel сделает всё сам)
            foreach (ExcelMiniDataBlock db in dataBlockList)
            {
                db.Write(false);
            }
            
        }
        
        // Получить каждую строку из группы представления
        void ProcessGroup(ExcelExportConfiguration exco, RowElementGroup group, ExcelMiniDataBlock dblock)
        {
            foreach (RowElementGroup.Item item in group.Items)
            {
                Dictionary<Guid, RowElementGroup.MergedCellValue> MergedCells = item.MergedCells;
                List<RowElement> MergedElements = item.MergedElements;

                var excelRow = new Dictionary<string, object>();

                foreach (ParameterDescriptor pd in exco.pColumns)
                {
                    excelRow[pd.Name] = exco.GetCellValue(pd, item);
                  
                }
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
