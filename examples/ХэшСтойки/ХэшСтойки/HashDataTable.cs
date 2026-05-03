using FBase.Tables;
using System;
using System.Data;
using System.Linq;
using TFlex.Model;

namespace TFAPIHash
{
    public class HashDataTable : MyDataTable
    {
        public Document document;
        public InternalDatabase database;
        public Chops chops;

        public HashDataTable()
        {
        }
        public HashDataTable(Document Document, InternalDatabase database)
        {
            this.document = Document;
            this.database = database;

            chops = new Chops(Document, database);

            this.TableName = "hashes";
            this.ExtendedProperties.Add("Caption", "Хэши");

            DataColumn id = AddLongColumn("id", new[] { flag.AutoIncrement, flag.Unique });
            DataColumn prow = AddLongColumn(chops.rowid.TableColumnName); //  new[] { flag.Unique }
            DataColumn Плоскость = AddStringColumn(chops.Плоскость.TableColumnName);
            DataColumn Этаж = AddIntegerColumn(chops.Этаж.TableColumnName);
            DataColumn Стойка = AddIntegerColumn(chops.Стойка.TableColumnName);
            DataColumn ЛевоеЗаполнение = AddStringColumn(chops.ЗаполнениеСлева.TableColumnName);
            DataColumn ПравоеЗаполнение = AddStringColumn(chops.ЗаполнениеСправа.TableColumnName);
            DataColumn Длина = AddDoubleColumn(chops.ВысотаСтойки.TableColumnName);
            DataColumn Позиция = AddIntegerColumn(chops.Позиция.TableColumnName);
            DataColumn hash = AddStringColumn(chops.hash.TableColumnName);

            DataColumn count = AddIntegerColumn("Count"); // Столбец для хранения количества одинаковых хэшей

            this.PrimaryKey = new DataColumn[] { id };
        }

        public void SaveChanges()
        {
            database.Document.BeginChanges("");
            // Получаем только измененные и новые строки
            DataTable dtChanges = this.GetChanges(DataRowState.Modified | DataRowState.Added); //  | DataRowState.Added

            if (dtChanges != null)
            {
                // Теперь dtChanges содержит только то, что нужно сохранить в БД
                foreach (DataRow row in dtChanges.Rows)
                {
                    // Логика сохранения...
                    int rid = ConvertToInteger(row[chops.rowid.TableColumnName]);
                    foreach (Chop c in chops.chops)
                    {
                        if (string.Compare(c.TableColumnName, chops.rowid.TableColumnName) == 0) { continue; }
                        WriteChop(row, rid, c);
                    }
                }
            }
            /////////////////////////////////
            // Получаем только измененные и новые строки
            dtChanges = this.GetChanges(DataRowState.Added);

            if (dtChanges != null)
            {
                // Теперь dtChanges содержит только то, что нужно сохранить в БД
                foreach (DataRow row in dtChanges.Rows)
                {
                    // Логика сохранения...
                    int rid = database.AppendRow();
                    foreach (Chop c in chops.chops)
                    {
                        if (string.Compare(c.TableColumnName, chops.rowid.TableColumnName) == 0) { continue; }
                        WriteChop(row, rid, c);
                    }
                }
            }
            database.Document.EndChanges();
            database.Document.Changed = true;
            database.Document.Save();
            database.Document.Close();
        }

        public void ReadRows(InternalDatabase database)
        {
            int RecordCount = database.GetRecordCount();
            //int ColumnCount = database.GetColumnCount();
            for (int r = 0; r < RecordCount; r++)
            {
                DataRow newRow = this.NewRow();
                foreach (Chop c in chops.chops)
                {
                    if (string.Compare(c.TableColumnName, chops.rowid.TableColumnName) == 0)
                    {
                        newRow[c.TableColumnName] = r;
                        continue;
                    }
                    ReadChop(newRow, r, c);
                }

                //string Плоскость = ConvertToString(newRow[chops.Плоскость.TableColumnName]);
                //int Этаж = ConvertToInteger(newRow[chops.Этаж.TableColumnName]);
                //int Стойка = ConvertToInteger(newRow[chops.Стойка.TableColumnName]);

                // 2. Ищем строку в таблице
                //DataRow existingRow = this.Rows.Find(new object[] { Плоскость, Этаж, Стойка });

                //if (existingRow == null)
                //{
                this.Rows.Add(newRow);
                //}


            }

            this.AcceptChanges();
        }

        public bool HashCompare(DataRow row, Chop chop, string hash)
        {
            string data1 = "";
            if (row[chop.TableColumnName] == DBNull.Value)
            {
                data1 = "";
            }
            else
            {
                data1 = Convert.ToString(row[chop.TableColumnName]);
            }

            bool result = true;
            if (string.Compare(data1, hash) != 0)
            {
                row[chop.TableColumnName] = hash;
                // write hash to a new row
                result = false;
            }
            return result;
        }
        public bool StringCompare(DataRow row, Chop chop)
        {
            string data1 = ConvertToString(row[chop.TableColumnName]);
            string data2 = chop.TextDocumentVariableValue();

            bool result = true;
            if (string.Compare(data1, data2) != 0)
            {
                row[chop.TableColumnName] = data2;
                result = false;
            }
            return result;
        }
        public bool IntegerCompare(DataRow row, Chop chop)
        {
            int data1 = ConvertToInteger(row[chop.TableColumnName]);
            int data2 = chop.IntegerDocumentVariableValue();
            bool result = true;
            if (data1 != data2)
            {
                row[chop.TableColumnName] = data2;
                result = false;
            }
            return result;
        }
        public bool RealCompare(DataRow row, Chop chop)
        {
            double data1 = ConvertToDouble(row[chop.TableColumnName]);
            double data2 = chop.RealDocumentVariableValue();
            bool result = true;
            if (data1 != data2)
            {
                row[chop.TableColumnName] = data2;
                result = false;
            }
            return result;
        }

        public string ConvertToString(object Object, string DefaultValue = "")
        {
            string result = DefaultValue;
            if (Object != DBNull.Value)
            {
                result = Convert.ToString(Object);
            }
            return result;
        }
        public double ConvertToDouble(object Object, double DefaultValue = -1)
        {
            double result = DefaultValue;
            if (Object != DBNull.Value)
            {
                result = Convert.ToDouble(Object);
            }
            return result;
        }
        public int ConvertToInteger(object Object, int DefaultValue = -1)
        {
            int result = DefaultValue;
            if (Object != DBNull.Value)
            {
                result = Convert.ToInt32(Object);
            }
            return result;
        }

        private void ReadChop(DataRow row, int rowIndex, Chop chop)
        {
            switch (Type.GetTypeCode(chop.DataType))
            {
                case TypeCode.Int32:
                    row[chop.TableColumnName] = database.GetIntValue(chop.InternalDatabaseColumnIndex, rowIndex);
                    break;
                case TypeCode.String:
                    row[chop.TableColumnName] = database.GetTextValue(chop.InternalDatabaseColumnIndex, rowIndex);
                    break;
                case TypeCode.Double:
                    row[chop.TableColumnName] = database.GetRealValue(chop.InternalDatabaseColumnIndex, rowIndex);
                    break;
            }
        }
        private void WriteChop(DataRow row, int rowIndex, Chop chop)
        {
            switch (Type.GetTypeCode(chop.DataType))
            {
                case TypeCode.Int32:
                    database.SetIntValue(chop.InternalDatabaseColumnIndex, rowIndex, ConvertToInteger(row[chop.TableColumnName]));
                    break;
                case TypeCode.String:
                    database.SetTextValue(chop.InternalDatabaseColumnIndex, rowIndex, ConvertToString(row[chop.TableColumnName]));
                    break;
                case TypeCode.Double:
                    database.SetRealValue(chop.InternalDatabaseColumnIndex, rowIndex, ConvertToDouble(row[chop.TableColumnName]));
                    break;
            }
        }

        public void Group()
        {
            // Группируем данные по трем критериям: Левое, Правое, Длина
            var baseGroups = this.AsEnumerable()
                .GroupBy(row => new
                {
                    L = row.Field<string>(chops.ЗаполнениеСлева.TableColumnName),
                    R = row.Field<string>(chops.ЗаполнениеСправа.TableColumnName),
                    Len = row.Field<double>(chops.ВысотаСтойки.TableColumnName) // Высота стойки
                })
                .Select(group => new
                {
                    Key = group.Key,
                    Rows = group.ToList() // Все строки, входящие в эту комбинацию L, R, Len
                });

            // 2. Пример обхода полученных групп
            foreach (var baseGroup in baseGroups)
            {
                int ВысотаСтойки = ConvertToInteger(baseGroup.Key.Len);
                string запЛев = ConvertToString(baseGroup.Key.L);
                string запПрав = ConvertToString(baseGroup.Key.R);


                // Находим максимальное значение в колонке "Позиция" среди строк текущей группы
                // Используем DefaultIfEmpty(0), чтобы не получить ошибку, если группа пуста или там одни NULL
                int maxInGroup = baseGroup.Rows
                    .Where(r => r[chops.Позиция.TableColumnName] != DBNull.Value)
                    .Select(r => Convert.ToInt32(r[chops.Позиция.TableColumnName]))
                    .DefaultIfEmpty(-1)
                    .Max();
                if (maxInGroup < 0) { maxInGroup = -1; }

                // Делим на группы по значению хэша
                var hashSubGroups = baseGroup.Rows
                        .GroupBy(r => r.Field<string>(chops.hash.TableColumnName))
                        .Select(group => new
                        {
                            Key = group.Key,
                            Rows = group.ToList()
                        });


                int count = hashSubGroups.Count();

                // Обход группы с одинаковыми хэшами
                foreach (var hashSubGroup in hashSubGroups)
                {
                    string hash = ConvertToString(hashSubGroup.Rows[0][chops.hash.TableColumnName]);

                    if (string.Compare(hash, "486d83f75052e80f095466d32485831fce716c76b1c3f75618c48baa26977f1e") == 0)
                    {
                        int wer = 1;
                    }


                    // Ищем максимальное значение позиции в группе с одинаковыми хэшами.
                    // Подразумевается, что в группе всего 2 разных значения: 0 для новых строк и проставленная ранее позиция
                    int maxInHashSubGroup = hashSubGroup.Rows
                    .Where(r => r[chops.Позиция.TableColumnName] != DBNull.Value)
                    .Select(r => Convert.ToInt32(r[chops.Позиция.TableColumnName]))
                    .DefaultIfEmpty(0)
                    .Max();

                    if (maxInHashSubGroup < 0)
                    {
                        // В группе нет строк, инициализированных параметром "Позиция"
                        // Назначаем новый номер
                        maxInGroup++;
                        maxInHashSubGroup = maxInGroup;
                    }

                    foreach (DataRow row in hashSubGroup.Rows)
                    {
                        if (ConvertToInteger(row[chops.Позиция.TableColumnName]) != maxInHashSubGroup)
                        {
                            row[chops.Позиция.TableColumnName] = maxInHashSubGroup;
                            if (row.RowState == DataRowState.Unchanged)
                            {
                                row.SetModified();
                            }
                        }
                        else
                        {
                            // Отменить запись в базу данных
                            row.AcceptChanges();
                        }

                    }
                }

            }
        }

        public void CheckColumns()
        {
            int newIndex = database.GetColumnCount();
            if (newIndex == 8) { return; }
            if (newIndex != 0) { throw new Exception("Неправильное количество столбцов в базе данных"); }
            database.Document.BeginChanges("");
            database.InsertColumn(newIndex, DatabaseColumnType.DBText, chops.Плоскость.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBReal, chops.Этаж.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBReal, chops.Стойка.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBText, chops.ЗаполнениеСлева.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBText, chops.ЗаполнениеСправа.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBReal, chops.ВысотаСтойки.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBReal, chops.Позиция.TableColumnName);
            newIndex++;
            database.InsertColumn(newIndex, DatabaseColumnType.DBText, chops.hash.TableColumnName);
            database.Document.EndChanges();
            database.Document.Changed = true;
            database.Document.Save();
            //DatabaseDocument.Close();
        }
        /*
        public void Write()
        {
            foreach (var group in calculatedGroups)
            {
                foreach (var row in group.Rows)
                {
                    double currentHash = row.Field<double>("Hash");
                    int versionNumber = group.HashVersions[currentHash];
                    int rowId = row.Field<int>("Id"); // ID строки в БД

                    string updateSql = "UPDATE Stands SET HashVersion = @v WHERE Id = @id";

                    using (var command = new SqlCommand(updateSql, connection, transaction))
                    {
                        command.Parameters.AddWithValue("@v", versionNumber);
                        command.Parameters.AddWithValue("@id", rowId);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
        */
    }
}







