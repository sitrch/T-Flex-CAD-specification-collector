using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.Model;
using TFlex.Model.Data.ProductStructure;

namespace TFAPIHash
{
    public class ProductStructureExportToExcel
    {
        static Document document = TFlex.Application.ActiveDocument;
        string path = Path.Combine(document.FilePath, "ProductStructure.xlsx");
        ProductStructure ps;

        public void ExportProductStructure()
        {
            var dataList = new List<Dictionary<string, object>>();

            // Получаем первый активный состав изделия
            ICollection<ProductStructure> pss = document.GetProductStructures();
            foreach (ProductStructure tps in pss)
            {
                if (tps == null) { continue; }
                if (tps.GetName(ModelObjectName.Name) == "m2spec")
                {
                    ps = tps;
                    break;
                }
            }
            if (ps == null) return;

            ProductStructureExcelExportOptions options = new ProductStructureExcelExportOptions();
            options.DecimalDelimiter = ProductStructureExportOptions.DecimalDelimiterType.Comma;
            options.ExportGroupHeaders = true;

            Scheme scheme = ps.GetScheme();
            Groupings groupings = scheme.Properties.Groupings;
            GroupingRules myRules = groupings.FirstOrDefault(g => g.Name == "Спецификация");
            options.FilePath = path;
            options.Silent = true;

            if (myRules != null)
            {
                Guid id = myRules.ID;
                options.GroupingUID = id;
            }

            
            
            ps.ExportToExcel(options);
        }
    }
}
