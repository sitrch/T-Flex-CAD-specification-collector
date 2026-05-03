using TFlex.Model;

namespace TFAPIHash
{
    /// <summary>
    /// Класс единичного фильтра
    /// Каждая группа переменных соответствует своей таблице, названной по имени группы.
    /// Переменные и фильтры в группе относятся к одной и той же таблице
    /// </summary>
    public class DBFilter
    {
        // Имя группы переменных
        public string VariableGroup { get { return variable.GroupName; } }

        // Ссылка на переменную документа
        public Variable variable;
        public string FilterVariableName { get { return variable.Name; } }
        public bool IsReal { get { return variable.IsReal; } }
        public bool IsText { get { return variable.IsText; } }
        public double RealValue { get { return variable.RealValue; } }
        public string TextValue { get { return variable.TextValue; } }

        public string Flag;
        public bool SkipFiltering;

        public DatabaseColumnType DBFilterType
        {
            get
            {
                if (IsReal) { return DatabaseColumnType.DBReal; }
                else if (IsText) { return DatabaseColumnType.DBText; }
                return (DatabaseColumnType.DBUnknown);
            }
        }

        public DBFilter(Variable variable, string Flag = "")
        {
            this.variable = variable;
            this.Flag = Flag;

            SkipFiltering = false;
            switch (Flag)
            {
                case "flag counter": SkipFiltering = true; break;
            }
        }

    }
}
