using Newtonsoft.Json.Linq;
using pdmrwordplugin.Functions;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.Utilities
{
    public static class ClsRefPub
    {
        public static List<ReferenceModel> GetReferencesFromDoc()
        {
            try
            {
                List<ReferenceModel> objlist = null;
                Word.Document odoc = Globals.ThisAddIn.Application.ActiveDocument;
                if (odoc == null) { return null; }
                Word.Range orange = ClsCommonUtils.GetReferenceRangebyBook();
                if (orange == null)
                    orange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                if (orange == null) { return null; }
                if (orange.Paragraphs.Count == 1) { MessageBox.Show(ClsMessages.REF_MESSAGE_1, ClsGlobals.PROJ_TITLE); }
                if (orange.Paragraphs.Count > 1)
                {
                    //if (MessageBox.Show(ClsMessages.REF_MESSAGE_2, ClsGlobals.PROJ_TITLE, MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    //{
                    objlist = new List<ReferenceModel>(Functions.ClsCommonUtils.GetReferenceDetails(odoc, orange));
                    if (objlist != null && objlist.Count > 0) { return objlist; }
                    else
                    {
                        MessageBox.Show("Error while getting the references", ClsGlobals.PROJ_TITLE, MessageBoxButton.OK, MessageBoxImage.Warning);
                        return null;
                    }
                    //}
                }
                return objlist;
            }
            catch { return null; }
        }

        public static Task<List<ReferenceModel>> IParseReferencebyExe(List<ReferenceModel> reflists)
        {
            return Task.Run(() => ParseReferencebyExe(reflists));
        }

        private static void RemoveTempFiles(string file1)
        {
            try
            {
                if (System.IO.File.Exists(file1)) { System.IO.File.Delete(file1); }
            }
            catch { }
        }

        public static List<ReferenceModel> ParseReferencebyExe(List<ReferenceModel> processreflists)
        {
            try
            {
                string outJson = System.IO.Path.GetTempPath() + "refwordprocessed.json";
                string reffilepath = System.IO.Path.GetTempPath() + @"refword.txt";
                RemoveTempFiles(outJson);
                RemoveTempFiles(reffilepath);
                List<ReferenceModel> references = new List<ReferenceModel>();
                string reftext = string.Join(Environment.NewLine, processreflists.Select(x => x.Reftext));
                System.IO.File.WriteAllText(reffilepath, reftext, Encoding.UTF8);
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    FileName = "cmd.exe",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    Arguments = @"/c " + ClsGlobals.REF_PARSER_PATH + ClsGlobals.REF_PARSER_EXE + " > \"" + outJson + "\""
                };
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
                //Read references
                if (System.IO.File.Exists(outJson))
                {
                    string jsonString = System.IO.File.ReadAllText(outJson);
                    var getResults = JArray.Parse(jsonString);
                    int jindex = -1;
                    foreach(var reference in processreflists)
                    {
                        jindex++;
                        references.Add(new ReferenceModel()
                        {
                            Reftext = reference.Reftext,
                            Refbookmark = reference.Refbookmark,
                            RefJSON = getResults[jindex].ToString(),
                            PIndex = reference.PIndex,
                            ReftextHtml = reference.ReftextHtml
                        });  
                    }
                }
                //Ends here
                return references;
            }
            catch { return null; }
        }

        public static void ParseReferencebyOnlinePub()
        {

        }
        public static void ParseReferencebyOnlineXref()
        {

        }
    }
}
