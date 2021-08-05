using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(WorkWithDocument);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(WorkWithDocument);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Word.Paragraphs paras = Doc.Paragraphs;

            for (int i = 0; i < paras.Count; i++)
            {
                //paras[i].Range.;
            }

            string title = "Title";
            System.Windows.Forms.MessageBox.Show(Doc.Paragraphs[1].Range.Text, title);
            Doc.Paragraphs[1].Range.InsertParagraphBefore();
            Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
        }

        #endregion
        private void WorkWithDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            try
            {
                Word.Range rng = Doc.Range(0, 0);
                rng.Text = "New Text";
                rng.Select();
            }
            catch (Exception ex)
            {
                // Handle exception if for some reason the document is not available.
            }
           

        }
    }
}
