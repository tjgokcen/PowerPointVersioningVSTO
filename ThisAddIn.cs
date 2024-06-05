using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace PowerPointVersioningVSTO
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationBeforeSave += new PowerPoint.EApplication_PresentationBeforeSaveEventHandler(Application_PresentationBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_PresentationBeforeSave(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            try
            {
                var properties = Pres.CustomDocumentProperties;
                string versionNumber = Pres.Tags["TempVersionNumber"];

                if (!string.IsNullOrEmpty(versionNumber))
                {
                    bool propExists = false;

                    foreach (DocumentProperty prop in properties)
                    {
                        if (prop.Name == "VersionNumber")
                        {
                            propExists = true;
                            prop.Value = versionNumber;
                            break;
                        }
                    }

                    if (!propExists)
                    {
                        properties.Add("VersionNumber", false, MsoDocProperties.msoPropertyTypeString, versionNumber);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error saving version number: " + ex.Message);
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
