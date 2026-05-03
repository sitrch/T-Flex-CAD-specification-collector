using System.Data;


namespace FBase.Tables
{
    public class MyDataTable : DataTable
    {
        public enum flag
        {
            AllowDBNull,
            AutoIncrement,
            Unique,
            NOP
        }

        public DataColumn AddDoubleColumn(string ColumnName, flag[] flags)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(double));
            ImplementFlags(column, flags);
            this.Columns.Add(column);
            return column;
        }
        public DataColumn AddDoubleColumn(string ColumnName)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(double));
            this.Columns.Add(column);
            return column;
        }

        public DataColumn AddLongColumn(string ColumnName, flag[] flags)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(long));
            ImplementFlags(column, flags);
            this.Columns.Add(column);
            return column;
        }
        public DataColumn AddLongColumn(string ColumnName)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(long));
            this.Columns.Add(column);
            return column;
        }

        public DataColumn AddIntegerColumn(string ColumnName, flag[] flags, int DefaultValue = 0)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(int));
            column.DefaultValue = DefaultValue;
            ImplementFlags(column, flags);
            this.Columns.Add(column);
            return column;
        }

        public DataColumn AddIntegerColumn(string ColumnName, int DefaultValue = 0)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(int));
            column.DefaultValue = DefaultValue;
            this.Columns.Add(column);
            return column;
        }

        public DataColumn AddStringColumn(string ColumnName, flag[] flags)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(string));
            ImplementFlags(column, flags);
            this.Columns.Add(column);
            return column;
        }

        public DataColumn AddStringColumn(string ColumnName)
        {
            DataColumn column = new DataColumn(ColumnName, typeof(string));
            this.Columns.Add(column);
            return column;
        }
        private void ImplementFlags(DataColumn column, flag[] flags)
        {
            foreach (flag f in flags)
            {
                switch (f)
                {
                    case flag.AllowDBNull: column.AllowDBNull = true; break;
                    case flag.AutoIncrement: column.AutoIncrement = true; break;
                    case flag.Unique: column.Unique = true; break;
                    default:
                        column.AllowDBNull = false;
                        column.AutoIncrement = false;
                        column.Unique = false;
                        break;
                }
            }
        }

    }

}
