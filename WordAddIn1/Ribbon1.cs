using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {

           
            string url = "http://tahrirchi.net/api/corrections/";
            using (var webClient = new WebClient())
            {
                string title = "Title";
                var pars = new NameValueCollection();
                pars.Add("text", "bolaa");
                var response = webClient.UploadValues(url, pars);
                string str = System.Text.Encoding.UTF8.GetString(response);
                MessageBox.Show(str, title);

                Console.WriteLine(str);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Document nativeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            

            nativeDoc.Paragraphs[1].Range.InsertParagraphAfter();
            nativeDoc.Paragraphs[1].Range.Text = "Called from Test Button";
            nativeDoc.Paragraphs[1].Range.Font.Underline = WdUnderline.wdUnderlineDash;

        }
    }
}
