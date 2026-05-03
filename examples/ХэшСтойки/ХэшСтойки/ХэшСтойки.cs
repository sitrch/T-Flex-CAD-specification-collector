using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using TFlex.Model;
//using TFlex.Model.Model3D;
//using System.Xml.Linq;
//using static TFAPI_?????????.Test;
//using static TFlex.Model.Units.StandardUnits;


namespace TFAPIHash
{
    public class HashWriter
    {
        public Document document = TFlex.Application.ActiveDocument;
        public Document DatabaseDocument;

        public DBFiltersBundle filters;

        string HashDatabaseFileName = "dbHash.grb";

        //List<string> HashVariablesNames = new List<string>() { "" };
        string HashInputVariableName = "$хэшПолный";
        string HashOutputVariableName = "$hash";


        public void run()

        {
            int hash = CalcHash();

            //return;

            DatabaseDocument = TFAPIHash.db.FindDBase(document, HashDatabaseFileName);

            DatabaseDocument.BeginChanges("");

            filters = new DBFiltersBundle(document, DatabaseDocument);
            filters.SaveVariablesValuesToDatabase();

            filters.WriteCounter();



            DatabaseDocument.EndChanges();
            DatabaseDocument.Changed = true;
            DatabaseDocument.Save();
            DatabaseDocument.Close();


            //internalDatabase.MarkChanged();
            //internalDatabase.Regenerate(true);
        }





        /// <summary>
        /// Присваивает порядковые номера отфильтрованным записям в базе данных
        /// в рамках группы, образованной фильтрами
        /// запись номера позиции производится в фильтр с флагом "SkipFiltering"
        /// </summary>
        /// <param name="rows"></param>

        public int CalcHash()
        {
            Variable hashInputVariable = document.FindVariable(HashInputVariableName);
            Variable hashOutputVariable = document.FindVariable(HashOutputVariableName);
            string hash = hashInputVariable.TextValue;

            //int h = GetFastHash(hash);

            string outputHash = GetSha256Hash(hash);

            document.BeginChanges("");

            hashOutputVariable.TextValue = outputHash;
            document.EndChanges();




            //MessageBox.Show($"{hash}\n{outputHash}");

            return 0;
        }



        public static int GetFastHash(string input)
        {
            if (input == null) return 0;

            uint hash = 2166136261; // FNV offset basis
            const uint prime = 16777619; // FNV prime

            foreach (char c in input)
            {
                hash ^= c;
                hash *= prime;
            }

            return (int)hash;
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
