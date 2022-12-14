using CiteProc;
using CiteProc.Data;
using pdmrwordplugin.Functions;
using pdmrwordplugin.Taskpane;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Controls;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new pdmrRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace pdmrwordplugin
{
    [ComVisible(true)]
    public class pdmrRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public pdmrRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("pdmrwordplugin.pdmrRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void BtnClick(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnsetrefsel":
                    SetReferenceArea();
                    break;
                case "btnprocessXref":
                    AddReferenceUI("Xref");
                    break;
                case "btnprocessref":
                    AddReferenceUI("Reference");
                    break;
                case "btnstylexref":
                    AddReferenceUI("stylexref");
                    break;
                default:
                    break;
            }            
        }

        public void SetReferenceArea()
        {            
            ClsCommonUtils.SetReferenceRangebyBook();
        }

        public void AddReferenceUI(string wtype)
        {
            Microsoft.Office.Tools.CustomTaskPane objPane;
            switch (wtype)
            {                
                case "Reference":
                    Usrtaskpane oCtrl = new Usrtaskpane();                    
                    objPane = Globals.ThisAddIn.CustomTaskPanes.Add(oCtrl, ClsGlobals.PROJ_TITLE);
                    objPane.Control.Dock = System.Windows.Forms.DockStyle.Fill;
                    objPane.Width = 400;
                    objPane.Visible = true;
                    break;
                case "Xref":
                    if (!ClsCommonUtils.CreateXrefStyles()) { 
                        return; }
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView;
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal;
                    UsrXrefTaskpane oCtrl1 = new UsrXrefTaskpane();                    
                    objPane = Globals.ThisAddIn.CustomTaskPanes.Add(oCtrl1, ClsGlobals.PROJ_TITLE);
                    objPane.Control.Dock = System.Windows.Forms.DockStyle.Fill;
                    objPane.Width = 400;
                    objPane.Visible = true;
                    break;
                case "stylexref":
                    UsrStyleXrefs oCtrl2 = new UsrStyleXrefs();
                    objPane = Globals.ThisAddIn.CustomTaskPanes.Add(oCtrl2, ClsGlobals.PROJ_TITLE);
                    objPane.Control.Dock = System.Windows.Forms.DockStyle.Fill;
                    objPane.Width = 400;
                    objPane.Visible = true;
                    break;
                default:
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
