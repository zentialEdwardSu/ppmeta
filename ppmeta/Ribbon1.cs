using System;
using System.IO;
using System.Reflection;
using Office = Microsoft.Office.Core;

namespace ppmeta
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        // use static to keep instances alive
        private static TextEditorForm textEditorForm;
        private static FormatListForm formatListForm;
        private static PlaceholderForm placeholderForm;

        public Ribbon1()
        {
        }

        public void OnOpenMainClick(Office.IRibbonControl control)
        {
            // check if the window is already open
            if (textEditorForm != null && !textEditorForm.IsDisposed)
            {
                // bring the existing window to the front
                textEditorForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
                textEditorForm.BringToFront();
                textEditorForm.Activate();
            }
            else
            {
                // new windows
                textEditorForm = new TextEditorForm();
                // when the window is closed, clean up the reference
                textEditorForm.FormClosed += (s, e) => textEditorForm = null;
                textEditorForm.Show();
            }
        }

        public void OnFormatsClick(Office.IRibbonControl control)
        {
            var ppt = Globals.ThisAddIn.Application.ActivePresentation;
            
            if (formatListForm != null && !formatListForm.IsDisposed)
            {
                formatListForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
                formatListForm.BringToFront();
                formatListForm.Activate();
            }
            else
            {
                formatListForm = new FormatListForm(ppt);
                formatListForm.FormClosed += (s, e) => formatListForm = null;
                formatListForm.Show();
            }
        }

        public void OnOptionsClick(Office.IRibbonControl control)
        {
            var configForm = new ConfigForm(ShareState.Config);
            configForm.ShowDialog();
        }

        public void OnPlaceHolderClick(Office.IRibbonControl control)
        {
            var ppt = Globals.ThisAddIn.Application.ActivePresentation;
            
            if (placeholderForm != null && !placeholderForm.IsDisposed)
            {
                placeholderForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
                placeholderForm.BringToFront();
                placeholderForm.Activate();
            }
            else
            {
                placeholderForm = new PlaceholderForm(ppt);
                placeholderForm.FormClosed += (s, e) => placeholderForm = null;
                placeholderForm.Show();
            }
        }




        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ppmeta.Ribbon1.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region 帮助器

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
