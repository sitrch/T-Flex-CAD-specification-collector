using System;
using System.Collections.Generic;
using TFlex.Model;

namespace TFAPIHash
{
    /// <summary>
    /// SELECT DATA1, DATA2, ... FROM TABLE WHERE FILTER1 = filter1.value AND FILTER2 = filter2.value AND ...
    /// Класс для работы с таблицей
    /// Переносит значения переменных в таблицу
    /// Имя таблицы соответствует группе переменных
    /// </summary>
    /// <param name="VariableGroup"></param>
    /// <param name="Database"></param>
    /// <param name="DataVariablesList">Коллекция переменных с данными в этой группе фмльтров</param>
    public class DBTableFilters
    {
        Document document;

        public List<DBFilter> FiltersList = new List<DBFilter>();
        public List<Variable> DataVariablesList = new List<Variable>();

        public string VariableGroup;
        private InternalDatabase Database;

        string NameColumnName = "name";
        string RealColumnName = "real";
        string TextColumnName = "text";

        private int nameColumnIndex;
        private int realColumnIndex;
        private int textColumnIndex;


        public DBTableFilters(Document Document, string VariableGroup, InternalDatabase Database)
        {
            this.document = Document;
            this.VariableGroup = VariableGroup;
            this.Database = Database;
            nameColumnIndex = GetColumnIndex(NameColumnName, DatabaseColumnType.DBText, 1);
            realColumnIndex = GetColumnIndex(RealColumnName, DatabaseColumnType.DBReal, 2);
            textColumnIndex = GetColumnIndex(TextColumnName, DatabaseColumnType.DBText, 3);

        }

        public void AddDataVariable(Variable variable)
        {
            DataVariablesList.Add(variable);
        }


        /// <summary>
        /// Проверка соответствия строки фильтрам
        /// Пропускает проверку по фильтрам, где включен пропуск фильтрования SkipFiltering
        /// </summary>
        /// <param name="Database"></param>
        /// <param name="RowIndex"></param>
        /// <returns></returns>
        public bool Check(int RowIndex)
        {
            foreach (DBFilter f in FiltersList)
            {
                if (f.SkipFiltering) { continue; }

                for (int i = 0; i < Database.GetColumnCount(); i++)
                {
                    if (Database.GetColumnName(i) == f.FilterVariableName)
                    {
                        if (f.IsReal)
                        {
                            if (Database.GetRealValue(i, RowIndex) != f.RealValue) return false;
                        }
                        else if (f.IsText)
                        {
                            if (Database.GetTextValue(i, RowIndex) != f.TextValue) return false;
                        }
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// Ищет строку с нужной комбинацией фильтров
        /// В случае обнаружения новой комбинации фильтров для переменной, 
        /// создаётся строка с переменной и этой комбинацией фильтров
        /// </summary>
        /// <param name="Database"></param>
        /// <param name="RowIndex"></param>
        public void WriteNewFiltersCombination(int RowIndex, Variable variable)
        {
            Database.SetTextValue(nameColumnIndex, RowIndex, variable.Name);

            foreach (DBFilter f in FiltersList)
            {
                if (f.Flag == "flag counter") { continue; }

                for (int i = 0; i < Database.GetColumnCount(); i++)
                {
                    if (Database.GetColumnName(i) == f.FilterVariableName)
                    {
                        if (f.IsReal)
                        {
                            Database.SetRealValue(i, RowIndex, f.RealValue);

                        }
                        else if (f.IsText)
                        {
                            Database.SetTextValue(i, RowIndex, f.TextValue);
                        }
                    }
                }
            }


            //Database.Document.EndChanges();
            //Database.Document.Changed = true;
            //Database.Document.Save();
            //Database.Document.Close();
        }

        /// <summary>
        /// Ищет индекс строки в базе данных по имени переменной
        /// Проверяет, что строка соответствует комбинации фильтров
        /// Если строки с таким именем нет, то добавляет новую строку, 
        /// записывает в эту строку комбинацию фильтров,
        /// и возвращает индекс этой строки.
        /// Если найдена строка с пустым именем, то использует её индекс
        /// 
        /// </summary>
        /// <param name="variable"></param>
        /// <returns></returns>
        public int GetRowIndex(Variable variable)
        {
            int result = -1;
            int empty = -1;
            // 
            for (int r = 0; r < Database.GetRecordCount(); r++)
            {
                string s = Database.GetTextValue(nameColumnIndex, r);
                if (string.Compare(s, "") == 0 && empty == -1) { empty = r; }
                if (string.Compare(s, variable.Name) == 0)
                {
                    if (Check(r))
                    {
                        //Найдена строка с искомым именем и комбинацией фильтров
                        result = r;
                        return result;
                    }
                }
            }
            if (empty != -1)
            {
                result = empty;
            }
            else
            {
                int rec = Database.GetRecordCount();
                result = Database.AppendRow();
            }

            WriteNewFiltersCombination(result, variable);

            return result;

        }


        /// <summary>
        /// Ищет индекс столбца в базе данных по имени переменной
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <param name="DataColumnType"
        /// <param name="NewColumnIndex"></param>
        /// <returns></returns>

        public int GetColumnIndex(string ColumnName)
        {
            for (int c = 0; c < Database.GetColumnCount(); c++)
            {
                if (string.Compare(Database.GetColumnName(c), ColumnName) == 0)
                {
                    return c;
                }
            }

            return (-1);
        }
        public int GetColumnIndex(string ColumnName, DatabaseColumnType DataColumnType, int NewColumnIndex = -1)
        {
            int foundindex = GetColumnIndex(ColumnName);
            if (foundindex >= 0)
            {
                return (foundindex);
            }
            // New column
            if (NewColumnIndex == -1) { NewColumnIndex = Database.GetColumnCount(); }

            return (Database.InsertColumn(NewColumnIndex, DataColumnType, ColumnName));
        }

        /// <summary>
        /// Проверка и добавление столбцов в базу данных в зависимости от переменных, помеченных как фильтр
        /// !!! Сначала нужно заполнить лист FiltersList!!!
        /// </summary>
        /// <param name="Database"></param>
        /// <exception cref="Exception"></exception>

        public void CheckFilterColumns()
        {
            int result = -1;
            bool found = false;
            foreach (DBFilter f in FiltersList)
            {
                found = false;
                for (int c = 0; c < Database.GetColumnCount(); c++)
                {
                    string s = Database.GetColumnName(c);

                    if (string.Compare(f.FilterVariableName, s) == 0)
                    {
                        DatabaseColumnType dbcolumntype = Database.GetColumnType(c);
                        if (f.DBFilterType == dbcolumntype)
                        {
                            found = true;
                            break;
                        }

                    }
                }
                // new column
                if (!found)
                {
                    int newIndex = Database.GetColumnCount();
                    Database.InsertColumn(newIndex, f.DBFilterType, f.FilterVariableName);
                }


            }
        }

        public int InsertColumn(string NewColumnName, DatabaseColumnType NewColumnType)
        {
            int newIndex = Database.GetColumnCount();
            int result = Database.InsertColumn(newIndex, NewColumnType, NewColumnName);

            return result;
        }

        /// <summary>
        /// Запись данных из переменных в базу данных 
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="Database"></param>       
        public void WriteDatabase()
        {
            CheckFilterColumns();

            foreach (Variable dataVariable in DataVariablesList)
            {
                ///int row = GetRowIndex(dataVariable);
                ///Для хэшей стоек
                ///
                int row = GetRowIndex(dataVariable);

                if (dataVariable.IsReal)
                {
                    double rv = Database.GetRealValue(realColumnIndex, row);
                    // Здесь заменяет значение, если новое значение отличается от старого
                    if (rv != dataVariable.RealValue)
                    {
                        Database.SetRealValue(realColumnIndex, row, dataVariable.RealValue);
                    }
                }
                else if (dataVariable.IsText)
                {
                    string sv = Database.GetTextValue(textColumnIndex, row);
                    if (!sv.Equals(dataVariable.TextValue))
                    {
                        Database.SetTextValue(textColumnIndex, row, dataVariable.TextValue);
                    }
                }

            }

        }

        /// <summary>
        /// Записывает в столбец с флагом "flag counter"
        /// номер вхождения данных из столбца с флагом "flag counted"
        /// </summary>       
        public void GetCounters()
        {
            List<CountersStore> cStore = new List<CountersStore>();
            DisableSkippedOnCounted();
            DBFilter CounterFilter = GetCounterFilter();
            int CounterColumnIndex = GetColumnIndex(CounterFilter.FilterVariableName);

            int rc = Database.GetRecordCount();

            for (int r = 0; r < Database.GetRecordCount(); r++)
            {
                if (Check(r))
                {
                    //Найдена строка с искомым именем и комбинацией фильтров
                    // Получаем значение counter
                    int counter = (int)Database.GetRealValue(CounterColumnIndex, r);
                    cStore.Add(new CountersStore(CounterColumnIndex, r, counter));
                }
            }

            WriteCounters(cStore);
            WriteCounterVariable();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cstore"></param>
        public void WriteCounters(List<CountersStore> cstore)
        {
            int maxPosition = 0;

            foreach (CountersStore cs in cstore)
            {
                if (cs.Counter > maxPosition) { maxPosition = cs.Counter; }
            }

            foreach (CountersStore cs in cstore)
            {
                if (cs.Counter != 0) { continue; }
                maxPosition++;
                Database.SetRealValue(cs.ColumnIndex, cs.RowIndex, maxPosition);
            }
        }

        public void WriteCounterVariable()
        {
            EnableSkippedOnCounted();
            DBFilter CounterFilter = GetCounterFilter();
            int CounterColumnIndex = GetColumnIndex(CounterFilter.FilterVariableName);

            int rc = Database.GetRecordCount();

            for (int r = 0; r < Database.GetRecordCount(); r++)
            {
                if (Check(r))
                {
                    //Найдена строка с искомым именем и комбинацией фильтров
                    // Получаем значение counter

                    int c = (int)Database.GetRealValue(CounterColumnIndex, r);
                    Variable rv = document.FindVariable(CounterFilter.FilterVariableName);

                    document.BeginChanges("");
                    if (rv != null) { rv.RealValue = c; }
                    document.EndChanges();

                }
            }
        }


        /// <summary>
        /// Включает пропуск столбца с флагом "flag counted"
        /// при фильтрации
        /// позволяет получить массив из запроса
        /// </summary>
        public void DisableSkippedOnCounted()
        {
            foreach (DBFilter f in FiltersList)
            {
                if (f.Flag == "flag counted")
                {
                    f.SkipFiltering = true;
                    break;
                }
            }
        }

        public void EnableSkippedOnCounted()
        {
            foreach (DBFilter f in FiltersList)
            {
                if (f.Flag == "flag counted")
                {
                    f.SkipFiltering = false;
                    break;
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DBFilter GetCounterFilter()
        {
            foreach (DBFilter f in FiltersList)
            {
                if (f.Flag == "flag counter")
                {
                    return (f);
                }
            }
            return (null);
        }
    }



    /// <summary>
    /// 
    /// </summary>
    public class CountersStore
    {
        public int ColumnIndex = 0;
        public int RowIndex = 0;
        public int Counter = 0;
        public CountersStore(int ColumnIndex, int RowIndex, int Counter)
        {
            this.Counter = Counter;
            this.RowIndex = RowIndex;
            this.ColumnIndex = ColumnIndex;
        }
    }
}
