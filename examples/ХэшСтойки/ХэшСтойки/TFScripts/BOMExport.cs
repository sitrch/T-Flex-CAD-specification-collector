using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;

namespace TFAPIHash
{
    public class BOMExport
    {
        public static string BOMObjectName;
        public static ArrayList format = new ArrayList();
        public static ArrayList zone = new ArrayList();
        public static ArrayList position = new ArrayList();
        public static ArrayList number = new ArrayList();
        public static ArrayList desc = new ArrayList();
        public static ArrayList amount = new ArrayList();
        public static ArrayList remarks = new ArrayList();
        public static ArrayList material = new ArrayList();
        public static ArrayList mass = new ArrayList();
        public static ArrayList price = new ArrayList();

        public static TFlex.CAD.ExcelExportUtils.ExcelDocument excelDocument;

        public static void run()
        {
            Document document = TFlex.Application.ActiveDocument;
            if (document == null)
                return;

            SaveFileDialog saveFileDlg = new SaveFileDialog();
            saveFileDlg.Title = "Экспорт данных спецификации в Excel";
            saveFileDlg.Filter = "Документы Excel (*.xlsx) (/ OpenXML)|*.xlsx|Документы Excel (*.xls) (/ MS Jet)|*.xls";
            saveFileDlg.FilterIndex = 1;
            if (saveFileDlg.ShowDialog() == DialogResult.OK)
            {
                string xlsFileName = saveFileDlg.FileName;
                document.BeginChanges("Экспорт данных спецификации в Microsoft Excel");
                foreach (Text textobj in document.Texts)
                {
                    if (textobj is BOMObject)
                    {
                        BOMObjectName = ((BOMObject)textobj).FriendlyName;
                        if (!((BOMObject)textobj).ReportFileLink.IsEmpty)
                        {
                            Document doc = TFlex.Application.OpenDocument(((BOMObject)textobj).ReportFileLink);
                            if (doc != null)
                            {
                                doc.BeginChanges("Экспорт данных спецификации связанного документа в Microsoft Excel");
                                ModelObject textobject = doc.GetObjectByID(((BOMObject)textobj).ReportID);
                                if (textobject is BOMObject)
                                {
                                    SetFieldsValues((BOMObject)textobject, xlsFileName);
                                }
                                doc.EndChanges();
                            }
                        }
                        else
                        {
                            SetFieldsValues((BOMObject)textobj, xlsFileName);
                        }
                    }
                }
                document.EndChanges();
            }
        }


        public static void SetFieldsValues(BOMObject bomobj, string xlsFileName)
        {
            bomobj.BeginEdit();

            format.Clear();
            zone.Clear();
            position.Clear();
            number.Clear();
            desc.Clear();
            amount.Clear();
            remarks.Clear();
            material.Clear();
            mass.Clear();
            price.Clear();

            if (bomobj.MoveToFrontRecord())
            {
                do
                {
                    format.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Format)));
                    zone.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Zone)));
                    position.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Position)));
                    number.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Number)));
                    desc.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Desc)));
                    amount.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Amount)));
                    remarks.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Remarks)));
                    material.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Material)));
                    mass.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Mass)));
                    price.Add(SetFieldFormat(bomobj.GetStandardFieldValue(BOMObject.StandardField.Price)));
                }
                while (bomobj.MoveToNextRecord());
            }
            bomobj.EndEdit();

            try
            {
                if (Path.GetExtension(xlsFileName) == ".xlsx")
                {
                    ArrayList data = new ArrayList();
                    data.Add(format);
                    data.Add(zone);
                    data.Add(position);
                    data.Add(number);
                    data.Add(desc);
                    data.Add(amount);
                    data.Add(remarks);
                    data.Add(material);
                    data.Add(mass);
                    data.Add(price);

                    string[] colNames = new string[] { "Формат", "Зона", "Позиция", "Обозначение", "Наименование", "Количество", "Комментарии", "Материал", "Масса", "Цена" };

                    string bomobjName = bomobj.FriendlyName;

                    if (excelDocument == null)
                    {
                        excelDocument = new TFlex.CAD.ExcelExportUtils.ExcelDocument(xlsFileName, data, colNames, bomobjName);
                    }
                    else
                    {
                        excelDocument.WriteData(xlsFileName, data, colNames, true, bomobjName);
                    }
                }
                else
                    CreateTableViaMSJet(bomobj, xlsFileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static string SetFieldFormat(string field)
        {
            string str = field.Replace("%%119", "°").Replace("%%d", "°").Replace("%%066", "Ø").Replace("%%c", "Ø").Replace("%%042", "х").Replace("%%p", "±").Replace("\r\n", " ").Replace("\n", " ").Replace("%%S", " ").Replace("%%", "");
            return str;
        }

        public static void CreateTableViaMSJet(BOMObject bomobj, string xlsFileName)
        {
            string tcrConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + xlsFileName + ";" + "Extended Properties=\"Excel 8.0;HDR=YES\"";

            using (OleDbConnection insertConn = new OleDbConnection(tcrConnStr))
            {
                insertConn.Open();

                string tableName = BOMObjectName;
                string create = "create table " + tableName.Replace(" ", "_") + " (Формат varchar, Зона varchar, Позиция varchar, Обозначение varchar, Наименование varchar, Количество varchar, Комментарии varchar, Материал varchar, Масса varchar, Цена varchar)";

                using (OleDbCommand createCmd = new OleDbCommand(create, insertConn))
                {
                    createCmd.ExecuteNonQuery();
                }

                OleDbCommand insertCmd = new OleDbCommand();
                insertCmd.Connection = insertConn;
                for (int i = 0; i < format.Count; ++i)
                {
                    string insert = "insert into " + tableName.Replace(" ", "_") + " (Формат, Зона, Позиция, Обозначение, Наименование, Количество, Комментарии, Материал, Масса, Цена) values ('" + format[i].ToString() + "', '" + zone[i].ToString() + "', '" + position[i].ToString() + "', '" + number[i].ToString() + "', '" + desc[i].ToString() + "', '" + amount[i].ToString() + "', '" + remarks[i].ToString() + "', '" + material[i].ToString() + "', '" + mass[i].ToString() + "', '" + price[i].ToString() + "')";
                    insertCmd.CommandText = insert;
                    insertCmd.ExecuteNonQuery();
                }
            }
        }
    }
}



