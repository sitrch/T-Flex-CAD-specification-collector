using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using TFlex.Model;
//using TFlex.Model.Model3D.Geometry;

namespace TFAPIHash
{
    public class HashStore
    {
        Document document = TFlex.Application.ActiveDocument;
        static Document DatabaseDocument;
        string DatabaseDocumentFileName = "dbHash.grb";

        HashDataTable hTable;

        InternalDatabase database;
        public string DatabaseName = "HashStore";

        Chops chops;

        public DBFiltersBundle filters;

        string ProjectDatabaseNameVariableName = "$HashDatabaseName";
        string hash;


        public string Плоскость { get { return (chops.Плоскость.TextDocumentVariableValue()); } }
        public int Этаж { get { return (chops.Этаж.IntegerDocumentVariableValue()); } }
        public int Стойка { get { return (chops.Стойка.IntegerDocumentVariableValue()); } }
        public string ЗаполнениеСлева { get { return (chops.ЗаполнениеСлева.TextDocumentVariableValue()); } }
        public string ЗаполнениеСправа { get { return (chops.ЗаполнениеСправа.TextDocumentVariableValue()); } }
        public double Длина { get { return (chops.ВысотаСтойки.RealDocumentVariableValue()); } }
        public string HashDocumentValue { get { return (chops.hash.TextDocumentVariableValue()); } }
        public int ПозицияDocumentValue { get { return (chops.Позиция.IntegerDocumentVariableValue()); } }


        public HashStore()
        {
            DatabaseDocument = db.FindDBase(document, DatabaseDocumentFileName);
            database = db.GetInternalDatabase(DatabaseDocument, DatabaseName);
            chops = new Chops(document, database);
            hTable = new HashDataTable(document, database);
        }

        /// <summary>
        /// Запускается макросом из сборки Стойка
        /// 
        /// </summary>
        public void Run()
        {
            hash = GetSha256Hash(HashDocumentValue);
            if (string.Compare(hash, "a3aa99d10b055ebdbdc0df69b1809be95ac5c50bef7ba614d760bba7d72db97f") == 0)
            {
                int wer = 1;
            }
            hTable.CheckColumns();
            hTable.ReadRows(database);
            WriteHashToDatabase(hash);
            hTable.Group();
            hTable.SaveChanges();
            document.Regenerate(new RegenerateOptions { Full = true });
            //document.Changed = true;
            document.Save();

        }

        public void WriteHashToDatabase(string hash)
        {
            bool appendAfter = false;

            var rowToUpdate = hTable.AsEnumerable()
                    .FirstOrDefault(r => r.Field<string>(chops.Плоскость.TableColumnName) == chops.Плоскость.TextDocumentVariableValue() &&
                         r.Field<int>(chops.Этаж.TableColumnName) == chops.Этаж.IntegerDocumentVariableValue() &&
                         r.Field<int>(chops.Стойка.TableColumnName) == chops.Стойка.IntegerDocumentVariableValue());

            if (rowToUpdate == null)
            {
                rowToUpdate = hTable.NewRow();
                appendAfter = true;
            }
            hTable.StringCompare(rowToUpdate, chops.Плоскость);
            hTable.IntegerCompare(rowToUpdate, chops.Этаж);
            hTable.IntegerCompare(rowToUpdate, chops.Стойка);

            hTable.StringCompare(rowToUpdate, chops.ЗаполнениеСлева);
            hTable.StringCompare(rowToUpdate, chops.ЗаполнениеСправа);
            hTable.RealCompare(rowToUpdate, chops.ВысотаСтойки);

            if (hTable.HashCompare(rowToUpdate, chops.hash, hash) == false) { rowToUpdate[chops.Позиция.TableColumnName] = -1; }

            if (appendAfter)
            {
                hTable.Rows.Add(rowToUpdate);
            }
        }

        public static string GetSha256Hash(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            using (SHA256 sha256 = SHA256.Create())
            {
                // 1. Преобразуем строку в массив байтов (UTF8 - стандарт)
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);

                // 2. Вычисляем хэш
                byte[] hashBytes = sha256.ComputeHash(inputBytes);

                // 3. Преобразуем байты в строку Hex (совместимо со всеми версиями .NET)
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }



    }
}
