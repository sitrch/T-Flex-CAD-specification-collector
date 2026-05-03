using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;

namespace Run
{
    public class BOMExportTable
    {
        public static TFlex.CAD.ExcelExportUtils.ExcelDocument excelDocument;

        public static void run()
        {
            Document document = TFlex.Application.ActiveDocument;
            if (document == null)
                return;

            SaveFileDialog saveFileDlg = new SaveFileDialog();
            saveFileDlg.Title = "Экспорт спецификации в Excel";
            saveFileDlg.Filter = "Документы Excel (*.xlsx) (/ OpenXML)|*.xlsx|Документы Excel (*.xls) (/ MS Jet)|*.xls";
            saveFileDlg.FilterIndex = 1;

            if (saveFileDlg.ShowDialog() != DialogResult.OK)
                return;

            string xlsFileName = saveFileDlg.FileName;
            document.BeginChanges("Экспорт спецификации в Microsoft Excel");

            foreach (TFlex.Model.Model2D.Text textobj in document.Texts)
            {
                if (textobj is BOMObject)
                {
                    string BomObjectName = ((BOMObject)textobj).FriendlyName;
                    if (!((BOMObject)textobj).ReportFileLink.IsEmpty)
                    {
                        Document doc = TFlex.Application.OpenDocument(((BOMObject)textobj).ReportFileLink);
                        if (doc != null)
                        {
                            doc.BeginChanges("Экспорт данных спецификации связанного документа в Microsoft Excel");
                            ModelObject textobject = doc.GetObjectByID(((BOMObject)textobj).ReportID);
                            if (textobject is BOMObject)
                                DoExport((BOMObject)textobject, xlsFileName, BomObjectName);
                            doc.EndChanges();
                        }
                    }
                    else
                    {
                        DoExport((BOMObject)textobj, xlsFileName, BomObjectName);
                    }
                }
            }
            document.EndChanges();
        }


        private static void DoExport(BOMObject bomObject, string xlsFileName, string bomObjectName)
        {
            try
            {
                if (Path.GetExtension(xlsFileName) == ".xlsx")
                    CreateTableViaOXML(bomObject, bomObjectName, xlsFileName);

                else
                    CreateTableViaMSJet(bomObject, bomObjectName, xlsFileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // экспорт через Open XML

        public static void CreateTableViaOXML(BOMObject bomObject, string bomObjectName, string xlsFileName)
        {
            bomObject.BeginEdit();
            string[] columnNames = bomObject.GetVisibleFields();
            bomObject.EndEdit();

            RichText text = (RichText)bomObject;
            text.BeginEdit();

            TFlex.Model.Model2D.Table table = text.GetTableByIndex(0);

            DataTable data = new DataTable();

            int nColumns = Math.Min(table.ColumnCount, columnNames.Length);

            for (int j = 0; j < nColumns; j++)
            {
                DataColumn column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = columnNames[j];
                column.DefaultValue = "";
                data.Columns.Add(column);
            }

            for (int i = 0; i < table.RowCount; i++)
            {
                DataRow row = data.NewRow();
                data.Rows.Add(row);
                int iHeadCell = i * table.ColumnCount;
                for (int j = 0; j < nColumns; j++)
                {
                    string colName = columnNames[j];
                    int iCell = iHeadCell + j;

                    table.SetCursorPosition((UInt32)iHeadCell, 0);
                    Position pos1 = new Position(0, 0, iCell),
                             pos2 = new Position(table.GetCellTextLength((UInt32)iCell) - 1, 0, iCell);

                    if (pos1 == pos2)
                        row[colName] = "";
                    else
                    {
                        string str = text.GetText(pos1, pos2);
                        row[colName] = str.Replace("%%119", "°").Replace("%%d", "°").Replace("%%066", "Ø").Replace("%%c", "Ø").Replace("%%042", "х").Replace("%%p", "±").Replace("\n", " ");
                    }
                }
            }

            text.EndEdit();

            if (excelDocument == null)
            {
                excelDocument = new TFlex.CAD.ExcelExportUtils.ExcelDocument(xlsFileName, data, bomObjectName);
            }
            else
            {
                excelDocument.WriteData(xlsFileName, data, true, bomObjectName);
            }
        }


        // экспорт через MS Jet

        public static void CreateTableViaMSJet(BOMObject bomObject, string bomObjectName, string xlsFileName)
        {
            ArrayList _collection = new ArrayList();
            RichText text = (RichText)bomObject;
            text.BeginEdit();
            Table table = text.GetTableByIndex(0);
            for (int i = 0; i < table.ColumnCount; i++)
            {
                _collection.Add(new ArrayList());
                for (int j = 0; j < table.RowCount; j++)
                {
                    table.SetCursorPosition(System.Convert.ToUInt32(j * table.ColumnCount), 0);
                    Position pos1 = new Position(0, 0, j * table.ColumnCount + i);
                    Position pos2 = new Position(table.GetCellTextLength(System.Convert.ToUInt32(j * table.ColumnCount + i)) - 1, 0, j * table.ColumnCount + i);
                    if (pos1 == pos2)
                        ((ArrayList)_collection[i]).Add("");
                    else
                    {
                        string str = text.GetText(pos1, pos2);
                        ((ArrayList)_collection[i]).Add(str.Replace("%%119", "°").Replace("%%d", "°").Replace("%%066", "Ø").Replace("%%c", "Ø").Replace("%%042", "х").Replace("%%p", "±").Replace("\n", " "));
                    }
                }
            }
            text.EndEdit();

            CreateTableViaMSJet(bomObject, bomObjectName, xlsFileName, _collection);
        }

        private static void CreateTableViaMSJet(BOMObject bomObject, string bomObjectName, string xlsFileName, ArrayList collection)
        {
            string insertConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + xlsFileName + ";" + "Extended Properties=\"Excel 8.0;HDR=YES\"";

            using (OleDbConnection insertConn = new OleDbConnection(insertConnStr))
            {
                insertConn.Open();

                string tableName = bomObjectName;
                bomObject.BeginEdit();
                string[] columns_names = bomObject.GetVisibleFields();
                bomObject.EndEdit();

                for (int i = 0; i < columns_names.Length; i++)
                {
                    columns_names[i] = columns_names[i].Replace(" ", "_");
                    for (int j = 0; j < columns_names[i].Length; j++)
                    {
                        if (!Char.IsLetter(columns_names[i][j]) && !Char.IsNumber(columns_names[i][j]) && columns_names[i][j] != '_')
                            columns_names[i] = columns_names[i].Replace(columns_names[i][j], '_');
                    }
                }

                string create = "create table " + tableName.Replace(" ", "_") + " (";

                for (int i = 0; i < columns_names.Length - 1; i++)
                {
                    create += columns_names[i] + " varchar, ";
                }
                create += columns_names[columns_names.Length - 1] + " varchar)";
                using (OleDbCommand createCmd = new OleDbCommand(create, insertConn))
                {
                    createCmd.ExecuteNonQuery();
                }

                OleDbCommand insertCmd = new OleDbCommand();
                insertCmd.Connection = insertConn;

                string insert = "insert into " + tableName.Replace(" ", "_") + " (";
                string insert_string = "";
                for (int i = 0; i < columns_names.Length - 1; i++)
                {
                    insert += columns_names[i] + ", ";
                }
                insert += columns_names[columns_names.Length - 1] + ") values('";

                for (int j = 0; j < ((ArrayList)collection[1]).Count; j++)
                {
                    insert_string = insert;
                    for (int i = 0; i < collection.Count - 1; i++)
                    {
                        insert_string += ((ArrayList)collection[i])[j].ToString() + "', '";
                    }
                    insert_string += ((ArrayList)collection[collection.Count - 1])[j].ToString() + "')";
                    insertCmd.CommandText = insert_string;
                    insertCmd.ExecuteNonQuery();
                }
            }
        }
    }
}




