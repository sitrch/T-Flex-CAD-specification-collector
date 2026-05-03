using TFlex.Model;

namespace TFAPIHash
{
    public class FragmentsMacroRunner
    {
        Document document;
        string MacroName;
        public FragmentsMacroRunner(Document document, string MacroName)
        {
            this.document = document;
            this.MacroName = MacroName;
        }


        /// <summary>
        /// Запускает макрос вложенного фрагмента
        /// </summary>
        /// 
        public void Run()
        {


            foreach (var fragment in document.GetFragments())
            {
                // Получаем документ вложенного фрагмента
                Document fragmentDoc = fragment.GetFragmentDocument(true); // true для открытия на чтение/редактирование


                //bool ro = fragmentDoc.IsReadOnly;

                if (fragmentDoc != null)
                {
                    RunMacroInDocument(fragmentDoc, MacroName);
                    fragment.MarkChanged();
                    //fragmentDoc.Regenerate(new RegenerateOptions { Full = true});
                    fragmentDoc.Changed = true;
                    //fragmentDoc.Save();
                    //fragmentDoc.Close();

                }
                //fragmentDoc.EndChanges();
            }
            //document.BeginChanges("");
            //document.EndChanges();

            ////doc.Changed = true;
            ////doc.Regenerate(new RegenerateOptions { Full = true });
        }

        public void RunMacroInDocument(Document doc, string macroFullName)
        {
            // Проверяем наличие макроса в проекте документа
            if (!doc.MacroExists(macroFullName))
            {
                return;
            }

            doc.RunMacro(macroFullName);

        }


    }
}
