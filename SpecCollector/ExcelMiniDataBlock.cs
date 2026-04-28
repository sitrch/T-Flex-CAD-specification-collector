using System.Collections.Generic;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;

namespace SpecCollector
{
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
                FastMode = true,
                EnableAutoWidth = true
            };
        }
       
        public void Write(bool OverwriteFile = false)
        {
            MiniExcel.Insert(path, DataList, sheetName: SheetName, configuration: config);
        }

        public void AddRow(Dictionary<string, object> excelRow)
        {
            DataList.Add(excelRow);
        }

        public void CreateEmpty()
        {
            var columns = new List<Dictionary<string, object>>();
            MiniExcel.SaveAs(path, columns, overwriteFile: true);
        }
    }
}