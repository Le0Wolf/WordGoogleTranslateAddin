using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new TranslateButton();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace WordGoogleTranslateAddin.GUI
{
    using System.Drawing;
    using System.Windows.Forms;

    [ComVisible(true)]
    public class TranslateButton : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Action OnTranslateButtonPressed { get; set; }

        public TranslateButton()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordGoogleTranslateAddin.GUI.TranslateButton.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "translateItButton":
                    return new Bitmap(Properties.Resources.Google_Translate_Icon);
            }

            return null;
        }

        public void OnTranslatePressed(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "translateItButton":
                    this.OnTranslateButtonPressed();
                    break;
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
