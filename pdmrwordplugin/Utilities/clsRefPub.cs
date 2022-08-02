using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.Utilities
{
    public static class ClsRefPub
    {   
        public static List<ReferencePostions> GetReferencesFromDoc()
        {
            try
            {
                List<ReferencePostions> objlist = null;
                Word.Document odoc = Globals.ThisAddIn.Application.ActiveDocument;
                if (odoc == null) { return null; }
                Word.Range orange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                if (orange == null) { return null; }
                if (orange.Paragraphs.Count == 1) { MessageBox.Show(ClsMessages.REF_MESSAGE_1, ClsGlobals.PROJ_TITLE); }
                if (orange.Paragraphs.Count > 1)
                {
                    if (MessageBox.Show(ClsMessages.REF_MESSAGE_2, ClsGlobals.PROJ_TITLE, MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        objlist = new List<ReferencePostions>(Functions.ClsCommonUtils.GetReferenceDetails(odoc, orange));
                        if (objlist != null && objlist.Count > 0) { return objlist; }
                        else
                        {
                            MessageBox.Show("Error while getting the references", ClsGlobals.PROJ_TITLE, MessageBoxButton.OK, MessageBoxImage.Warning);
                            return null;
                        }
                    }
                }
                return objlist;
            }
            catch { return null; }
        }

        public static Task IParseReferencebyExe(List<ReferencePostions> reflists)
        {
            return Task.Run(() => ParseReferencebyExe(reflists));
        }
        public static void ParseReferencebyExe(List<ReferencePostions> processreflists)
        {

        }
        public static void ParseReferencebyOnlinePub()
        {

        }
        public static void ParseReferencebyOnlineXref()
        {

        }
    }
}
