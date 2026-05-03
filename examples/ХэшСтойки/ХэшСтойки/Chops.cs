using System.Collections.Generic;
using TFlex.Model;

namespace TFAPIHash
{
    public class Chops
    {
        public Document document;
        public InternalDatabase database;

        public List<Chop> chops = new List<Chop>();

        public Chop rowid;

        public Chop Плоскость;
        public Chop Этаж;
        public Chop Стойка;

        public Chop ЗаполнениеСлева;
        public Chop ЗаполнениеСправа;
        public Chop ВысотаСтойки;

        public Chop Позиция;
        public Chop hash;

        public Chops(Document Document, InternalDatabase database)
        {
            this.document = Document;
            this.database = database;
            rowid = new Chop(document, database, "", "rowid", (long)0);

            Плоскость = new Chop(document, database, "$Плоскость", "Плоскость", (string)"");
            Этаж = new Chop(document, database, "Этаж", "Этаж", (int)0);
            Стойка = new Chop(document, database, "Стойка", "Стойка", (int)0);

            ЗаполнениеСлева = new Chop(document, database, "$ЗаполнениеСлева", "ЗаполнениеСлева", (string)"");
            ЗаполнениеСправа = new Chop(document, database, "$ЗаполнениеСправа", "ЗаполнениеСправа", (string)"");
            ВысотаСтойки = new Chop(document, database, "ВысотаСтойки", "ВысотаСтойки", (double)0);

            Позиция = new Chop(document, database, "Позиция", "Позиция", (int)0);
            hash = new Chop(document, database, "$хэшПолный", "hash", (string)"");

            chops.Add(rowid);
            chops.Add(Плоскость); chops.Add(Этаж); chops.Add(Стойка);
            chops.Add(ЗаполнениеСлева); chops.Add(ЗаполнениеСправа); chops.Add(ВысотаСтойки);
            chops.Add(Позиция); chops.Add(hash);
        }



    }
}
