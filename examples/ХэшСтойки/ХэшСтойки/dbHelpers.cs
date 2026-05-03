using System;
using System.Data;
using System.IO;
using TFlex.Model;
//using static TFlex.Model.Model3D.LoftOperation;

namespace TFAPIHash
{
    public static class db
    {
        public static int InsertColumn(InternalDatabase database, string NewColumnName, DatabaseColumnType NewColumnType)
        {
            int newIndex = database.GetColumnCount();
            int result = database.InsertColumn(newIndex, NewColumnType, NewColumnName);

            return result;
        }

        public static int FindColumnIndex(InternalDatabase database, string ColumnName)
        {
            int ColumnCount = database.GetColumnCount();
            for (int i = 0; i < ColumnCount; i++)
            {
                string s = database.GetColumnName(i);
                if (string.Compare(s, ColumnName) == 0)
                {
                    return i;
                }

            }
            throw new Exception($"Не найден столбец: {ColumnName}");
        }



        /// <summary>
        /// Получает базу данных по имени
        /// Если таблица не существует, то
        /// создаёт таблицу методами API.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="db_name"></param>
        /// <returns></returns>
        public static InternalDatabase GetInternalDatabase(Document document, string db_name)
        {
            foreach (var i in document.GetDatabases())
            {
                if (i.Name == db_name && i.SubType == DatabaseType.Internal)
                {
                    return (InternalDatabase)i;
                }
            }
            document.BeginChanges("");
            InternalDatabase idb = new InternalDatabase(document, db_name);
            document.EndChanges();
            document.Save();
            return idb;
        }

        /// <summary>
        /// Ищет базу данных в родительских директориях
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ProjectDatabaseNameVariableName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static Document FindDBase(Document document, string dbFileName)
        {
            Document dbDoc = null;
            string startPath = document.FilePath;
            DirectoryInfo current = new DirectoryInfo(startPath);
            DirectoryInfo root = null;

            // 1. Ищем родительскую папку "Амбер" вверх по дереву
            while (current != null)
            {
                if (current.Name.Equals("Амбер", StringComparison.OrdinalIgnoreCase))
                {
                    root = current;
                    break;
                }
                current = current.Parent;
            }

            if (root != null)
            {
                // Получаем массив объектов FileInfo вместо строк

                FileInfo[] files = root.GetFiles(dbFileName, SearchOption.AllDirectories);

                foreach (FileInfo f in files)
                {
                    // f.Directory.Name — это имя папки, в которой лежит файл
                    if (f.Directory.Name.Equals("Data", StringComparison.OrdinalIgnoreCase))
                    {
                        dbDoc = TFlex.Application.OpenDocument(f.FullName, false, false);
                        if (dbDoc == null)
                        {
                            throw new Exception("Ошибка открытия документа с базой данных"); // todo
                        }

                        break;
                    }
                }
            }
            else
            {
                throw new Exception("База данных не найдена.");
            }
            return (dbDoc);
        }

        public static void WriteDatabase(InternalDatabase database, int ColumnIndex, int RowIndex, double value)
        {
            double rv = database.GetRealValue(ColumnIndex, RowIndex);
            // Здесь заменяет значение, если новое значение отличается от старого
            if (rv != value)
            {
                database.SetRealValue(ColumnIndex, RowIndex, value);
            }
        }
        public static void WriteDatabase(InternalDatabase database, int ColumnIndex, int RowIndex, string value)
        {
            string rv = database.GetTextValue(ColumnIndex, RowIndex);
            // Здесь заменяет значение, если новое значение отличается от старого
            if (string.Compare(rv, value) != 0)
            {
                database.SetTextValue(ColumnIndex, RowIndex, value);
            }
        }

        public static string GetTextDocumentVariableValue(Document doc, string DocumentVariableName)
        {
            Variable var = GetDocumentVariable(doc, DocumentVariableName);
            return (var.TextValue);
        }
        public static int GetIntegerDocumentVariableValue(Document doc, string DocumentVariableName)
        {
            Variable var = GetDocumentVariable(doc, DocumentVariableName);
            return ((int)var.RealValue);
        }

        public static void SetTextDocumentVariableValue(Document doc, string DocumentVariableName, string value)
        {
            Variable var = GetDocumentVariable(doc, DocumentVariableName);
            var.TextValue = value;
        }
        public static void SetIntegerDocumentVariableValue(Document doc, string DocumentVariableName, int value)
        {
            Variable var = GetDocumentVariable(doc, DocumentVariableName);
            var.RealValue = value;
        }
        public static Variable GetDocumentVariable(Document doc, string DocumentVariableName)
        {
            Variable var = doc.FindVariable(DocumentVariableName);
            if (var == null) { throw new Exception($"Не найдена переменная {DocumentVariableName}"); }
            return var;
        }

        /// <summary>
        /// Отбражает форму с таблицей
        /// Выбирает данные из InternalDatabase to DataTable
        /// </summary>
        /// <param name="internalDatabase"></param>
        /// <returns></returns>
        static DataTable GetDataTable(InternalDatabase internalDatabase)
        {
            // Отбражает форму с таблицей
            //dataTable = GetDataTable(internalDatabase);
            //Form1 form = (Form1)Activator.CreateInstance(typeof(Form1));
            // form.ShowData(dataTable);
            //form.ShowData(dataTable);
            //form.ShowDialog();
            DataTable table = null;
            DataColumn column = null;
            DataRow row = null;

            if (internalDatabase == null)
                return table;

            table = new DataTable(internalDatabase.Name);

            for (int i = 0; i < internalDatabase.GetColumnCount(); i++)
            {
                switch (internalDatabase.GetColumnType(i))
                {
                    case DatabaseColumnType.DBInt:
                        column = new DataColumn(internalDatabase.GetColumnName(i), typeof(Int32));
                        table.Columns.Add(column);
                        break;
                    case DatabaseColumnType.DBReal:
                        column = new DataColumn(internalDatabase.GetColumnName(i), typeof(Double));
                        table.Columns.Add(column);
                        break;
                    case DatabaseColumnType.DBText:
                        column = new DataColumn(internalDatabase.GetColumnName(i), typeof(String));
                        table.Columns.Add(column);
                        break;
                }
            }

            for (int r = 0; r < internalDatabase.GetRecordCount(); r++)
            {
                row = table.NewRow();

                for (int c = 0; c < internalDatabase.GetColumnCount(); c++)
                {
                    switch (internalDatabase.GetColumnType(c))
                    {
                        case DatabaseColumnType.DBInt:
                            row[c] = internalDatabase.GetIntValue(c, r);
                            break;
                        case DatabaseColumnType.DBReal:
                            row[c] = internalDatabase.GetRealValue(c, r);
                            break;
                        case DatabaseColumnType.DBText:
                            row[c] = internalDatabase.GetTextValue(c, r);
                            break;
                    }
                }

                table.Rows.Add(row);
            }

            return table;
        }

    }
}
