using System;
using TFlex.Model;

namespace TFAPIHash
{
    public class Chop
    {
        private InternalDatabase database;
        private Document document;

        public string DocumentVariableName;
        public string TableColumnName;
        public Type DataType;
        public DatabaseColumnType InternalDataBaseColumnType;
        public int InternalDatabaseColumnIndex { get { return (GetInternalDatabaseColumnIndex()); } }
        int _InternalDatabaseColumnIndex = -1;
        public string StringValue;
        public double DoubleValue;
        public int IntegerValue;

        public Chop(Document document, InternalDatabase Database, string DocumentVariableName, string TableColumnName, string Value)
        {
            this.database = Database;
            this.document = document;
            this.DocumentVariableName = DocumentVariableName;
            this.TableColumnName = TableColumnName;
            this.DataType = typeof(string);
            this.InternalDataBaseColumnType = DatabaseColumnType.DBText;
            this.StringValue = Value;
        }
        public Chop(Document document, InternalDatabase Database, string DocumentVariableName, string TableColumnName, double Value)
        {
            this.database = Database;
            this.document = document;
            this.DocumentVariableName = DocumentVariableName;
            this.TableColumnName = TableColumnName;
            this.DataType = typeof(double);
            this.InternalDataBaseColumnType = DatabaseColumnType.DBReal;
            this.DoubleValue = Value;
        }
        public Chop(Document document, InternalDatabase Database, string DocumentVariableName, string TableColumnName, int Value)
        {
            this.database = Database;
            this.document = document;
            this.DocumentVariableName = DocumentVariableName;
            this.TableColumnName = TableColumnName;
            this.DataType = typeof(int);
            this.InternalDataBaseColumnType = DatabaseColumnType.DBInt;
            this.IntegerValue = Value;
        }

        public string TextDocumentVariableValue()
        {
            Variable var = GetDocumentVariable();
            return (var.TextValue);
        }
        public int IntegerDocumentVariableValue()
        {
            Variable var = GetDocumentVariable();
            return ((int)var.RealValue);
        }
        public double RealDocumentVariableValue()
        {
            Variable var = GetDocumentVariable();
            return (var.RealValue);
        }

        public Variable GetDocumentVariable()
        {
            Variable var = document.FindVariable(DocumentVariableName);
            if (var == null) { throw new Exception($"Не найдена переменная {DocumentVariableName}"); }
            return var;
        }

        private int GetInternalDatabaseColumnIndex()
        {
            if (_InternalDatabaseColumnIndex < 0)
            { _InternalDatabaseColumnIndex = db.FindColumnIndex(database, TableColumnName); }
            return _InternalDatabaseColumnIndex;
        }
    }
}
