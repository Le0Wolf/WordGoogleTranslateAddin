using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordGoogleTranslateAddin
{
    using System.IO;
    using System.Windows.Forms;

    using Google.Apis.Auth.OAuth2;
    using Google.Cloud.Translation.V2;

    using WordGoogleTranslateAddin.GUI;

    public partial class ThisAddIn
    {
        private TranslationClient translationClient;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            using (var key = new MemoryStream(Properties.Resources.GoogleCloudPrivateKey))
            {
                var credential = GoogleCredential.FromStream(key);
                this.translationClient = TranslationClient.Create(credential, TranslationModel.NeuralMachineTranslation);
            }
        }

        private void TranslateSelected()
        {
            var selected = this.Application.Selection;
            if (selected.Type == Word.WdSelectionType.wdNoSelection)
            {
                return;
            }

            var text = selected.Text;
            var formatedText = selected.FormattedText.WordOpenXML;

            var translationResult = this.translationClient.TranslateText(selected.Text, "ru", "en");
            selected.Text = translationResult.TranslatedText;

            //var selectedText = Clipboard.GetData(DataFormats.Html);
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new TranslateButton
            {
                OnTranslateButtonPressed = this.TranslateSelected
            };
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
