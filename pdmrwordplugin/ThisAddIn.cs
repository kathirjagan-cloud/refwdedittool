using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.Management.Automation;
using System.Xml.Serialization;

namespace pdmrwordplugin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Add code to remove 
            ClsGlobals.APP_PATH = Environment.ExpandEnvironmentVariables(ClsGlobals.APP_PATH);
            if (File.Exists(ClsGlobals.APP_PATH + ClsGlobals.REF_STYLES_CONFIG))
            {
                ClsGlobals.gReferencestyles = Functions.ClsCommonUtils.DeserializeXMLFileToObject<Models.referencestyles>(ClsGlobals.APP_PATH + ClsGlobals.REF_STYLES_CONFIG);
            }
            // Ends here
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }
        
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new pdmrRibbon();
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
