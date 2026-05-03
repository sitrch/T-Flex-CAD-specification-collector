using System.Collections.Generic;
using TFlex.Model;

namespace TFAPIHash
{
    /// <summary>
    /// Содержит List DBFilter
    /// Количество записей соответствует количеству групп переменных
    /// </summary>
    public class DBFiltersBundle
    {
        private List<DBTableFilters> DBFiltersBundleList = new List<DBTableFilters>();
        private List<string> VariableGroupsList = new List<string>();
        private Document document;
        private Document databaseDocument;
        string VariableComment = "filter";
        string DataComment = "data";
        string FlagComment = "flag";
        public DBFiltersBundle(Document document, Document DatabaseDocument)
        {
            this.document = document;
            this.databaseDocument = DatabaseDocument;

            FillDBFilters();

        }



        /// <summary>
        /// Ищет подходящую переменную
        /// </summary>
        /// <param name="Document"></param>
        /// <param name="VariableComment">Идентификатор переменных-фильтров</param>
        private void FillDBFilters()
        {
            //DBFiltersBundleList.Clear();
            foreach (Variable v in document.GetVariables())
            {
                string s = v.Name;
                string comment = v.Comment;

                if (v.Hidden && !v.Visible) { continue; }
                if (v.Comment == VariableComment)
                {
                    AddToFilters(v);
                }
                if (v.Comment.StartsWith(FlagComment))
                {
                    AddToFilters(v, v.Comment);
                }
                else if (v.Comment == DataComment)
                {
                    AddToData(v);
                }
                else { continue; }
            }
        }

        /// <summary>
        /// Распределяет переменные по группам в зависимости от группы переменных в документе
        /// Создайт новую группу (new filter group), если её нет
        /// </summary>
        /// <param name="DBFilter"></param>
        private void AddToFilters(Variable variable, string Flag = "")
        {
            foreach (DBTableFilters f in DBFiltersBundleList)
            {
                if (f.VariableGroup == variable.GroupName)
                {
                    f.FiltersList.Add(new DBFilter(variable, Flag));
                    return;
                }
            }
            // new filter group
            CreateNewFiltersGroup(variable, Flag);
        }

        private void AddToData(Variable variable)
        {
            foreach (DBTableFilters f in DBFiltersBundleList)
            {
                if (f.VariableGroup == variable.GroupName)
                {
                    f.AddDataVariable(variable);
                    return;
                }
            }
            // new filter group
            InternalDatabase db = GetInternalDatabase(databaseDocument, variable.GroupName);
            DBTableFilters dbf = new DBTableFilters(document, variable.GroupName, db);
            dbf.DataVariablesList.Add(variable);
            DBFiltersBundleList.Add(dbf);
        }

        private void CreateNewFiltersGroup(Variable variable, string Flag = "")
        {
            InternalDatabase db = GetInternalDatabase(databaseDocument, variable.GroupName);
            DBTableFilters dbf = new DBTableFilters(document, variable.GroupName, db);
            dbf.FiltersList.Add(new DBFilter(variable, Flag));
            DBFiltersBundleList.Add(dbf);
        }

        /// <summary>
        /// Получает базу данных по имени
        /// Если таблица не существует, то
        /// создаёт таблицу методами API.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="db_name"></param>
        /// <returns></returns>
        public InternalDatabase GetInternalDatabase(Document document, string db_name)
        {
            foreach (var i in document.GetDatabases())
            {
                if (i.Name == db_name && i.SubType == DatabaseType.Internal)
                {
                    return (InternalDatabase)i;
                }
            }

            return new InternalDatabase(document, db_name);
        }

        /// <summary>
        /// Проверяет соответствие строк в таблице с переменными-фильтрами
        /// Проверяет наличие базовых строк для значений переменных
        /// </summary>

        public void SaveVariablesValuesToDatabase()
        {
            foreach (DBTableFilters f in DBFiltersBundleList)
            {
                f.WriteDatabase();
            }
        }

        /// <summary>
        /// Записывает в столбец с флагом "flag counter"
        /// номер вхождения данных из столбца с флагом "flag counted"
        /// </summary>
        public void WriteCounter()
        {
            foreach (DBTableFilters f in DBFiltersBundleList)
            {
                f.GetCounters();
            }
        }

    }
}
