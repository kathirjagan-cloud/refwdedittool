using pdmrwordplugin.Functions;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media.Animation;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.Utilities
{
    public static class clsVanXref
    {
        public static Task<List<XrefModel>> ReadCitationsfromDoc()
        {
            return Task.Run(() =>
            {
                return FindCitationsinDoc();
            });
        }

        public static List<XrefModel> FindCitationsinDoc()
        {
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                List<XrefModel> tmplist = new List<XrefModel>();
                int booktostart = 0;
                var list = FindCrossRefs("", "Super", booktostart, true);
                if (list != null && list.Count > 0)
                {
                    int.TryParse(list.LastOrDefault().XrefBookmark.Replace("XREF_", ""), out booktostart);
                    tmplist.AddRange(list);
                }
                list = FindCrossRefs(@"(\\[)(([0-9]+)([ ,\\;\\-\\–]+)?)+(\\])", "Square", booktostart, false);
                if (list != null && list.Count > 0)
                {
                    int.TryParse(list.LastOrDefault().XrefBookmark.Replace("XREF_", ""), out booktostart);
                    tmplist.AddRange(list);
                }
                list = FindCrossRefs(@"(\()(([0-9]+)([ ,\;\-\–]+)?)+(\))", "Round", booktostart, false);
                if (list != null && list.Count > 0)
                {
                    int.TryParse(list.LastOrDefault().XrefBookmark.Replace("XREF_", ""), out booktostart);
                    tmplist.AddRange(list);
                }
                list = FindCrossRefs(@"[rR]ef([eference(s)?])?([.: ]+)?(([\[\(])?(([0-9]+)([, \-\–]+)?)+([\]\)])?([, \-\–]+)?)+", "Ref", booktostart, false);
                if (list != null && list.Count > 0)
                {
                    int.TryParse(list.LastOrDefault().XrefBookmark.Replace("XREF_", ""), out booktostart);
                    tmplist.AddRange(list);
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                return tmplist;
            }
            catch { Globals.ThisAddIn.Application.ScreenUpdating = true; return null; }
        }

        public static List<XrefModel> FindCrossRefs(string rgxPatrn, string strType, int bookstart, bool Suprscr = false)
        {
            try
            {
                Word.Document actdoc = Globals.ThisAddIn.Application.ActiveDocument;
                List<XrefModel> fndrefs = new List<XrefModel>();
                Word.Range ignrRng = null;
                int bookID = bookstart;
                if (actdoc.Bookmarks.Exists(ClsGlobals.REF_BOOK_NAME))
                    ignrRng = actdoc.Bookmarks[ClsGlobals.REF_BOOK_NAME].Range.Duplicate;
                if (ignrRng == null)
                {
                    MessageBox.Show(ClsMessages.REF_MESSAGE_5, ClsGlobals.PROJ_TITLE);
                    return null;
                }
                Regex orgx = new Regex(rgxPatrn, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                string alreadyfind = "";
                Word.Range prvsrng = null;
                Word.Range subprvsrng = null;
                foreach (Word.Range sRng in actdoc.StoryRanges)
                {
                    if (rgxPatrn != "")
                    {
                        if (orgx.IsMatch(sRng.Text))
                        {
                            foreach (Match m in orgx.Matches(sRng.Text))
                            {
                                if (alreadyfind.Contains("**" + m.Value + "**"))
                                {
                                    alreadyfind += "**" + m.Value + "**";
                                    Word.Range orng = sRng.Duplicate;
                                    orng.Find.ClearFormatting();
                                    string sfndtxt = m.Value.Trim();
                                    string subfndtxt = null;
                                    while (orng.Find.Execute(sfndtxt))
                                    {
                                        if (prvsrng != null)
                                        {
                                            if (orng.InRange(prvsrng)) { break; }
                                        }
                                        if (!orng.InRange(ignrRng) && orng.Font.Superscript == 0)
                                        {
                                            Word.Range subrng = orng.Duplicate;
                                            subrng.Find.ClearFormatting();
                                            if (strType == "ref")
                                            {
                                                subfndtxt = m.Captures[2].Value;
                                            }
                                            else
                                            {
                                                string stmp = m.Value;
                                                stmp = stmp.Replace(m.Captures[0].Value, "");
                                                stmp = stmp.Replace(m.Captures[m.Captures.Count - 1].Value, "");
                                                subfndtxt = stmp;
                                            }
                                            subfndtxt = subfndtxt.Trim();
                                            subrng.Find.Text = subfndtxt;
                                            subprvsrng = null;
                                            while (subrng.Find.Execute(subfndtxt))
                                            {
                                                if (subrng.InRange(orng)) { break; }
                                                if (subprvsrng != null && subrng.InRange(subprvsrng)) { break; }
                                                if (string.IsNullOrEmpty(subprvsrng.Text.Trim())) { break; }
                                                while (subrng.Text.EndsWith("\r") ||
                                                        subrng.Text.EndsWith("\f") ||
                                                        subrng.Text.EndsWith("\v") ||
                                                        subrng.Text.EndsWith(",") ||
                                                        subrng.Text.EndsWith(";") ||
                                                        subrng.Text.EndsWith("."))
                                                {
                                                    subrng.SetRange(subrng.Start, subrng.End - 1);
                                                }
                                                bookID++;
                                                fndrefs.Add(new XrefModel()
                                                {
                                                    XrefText = subrng.Text,
                                                    XrefType = strType,
                                                    XrefBookmark = "XREF_" + bookID,
                                                    XrefContext = ClsCommonUtils.GetContextfromRng(subrng)
                                                });
                                                actdoc.Bookmarks.Add("XREF_" + bookID, subrng.Duplicate);
                                                subprvsrng = subrng.Duplicate;
                                                subrng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                            }
                                        }
                                        prvsrng = orng.Duplicate;
                                        orng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        bookID = bookstart;
                        Word.Range orng = sRng.Duplicate;
                        orng.Find.ClearFormatting();
                        orng.Find.Text = "";
                        orng.Find.Font.Superscript = 1;
                        while (orng.Find.Execute())
                        {
                            if (prvsrng != null && orng.InRange(prvsrng)) { break; }
                            if (orng.InRange(ignrRng)) { break; }
                            //orng.Select();
                            Word.Range tmporng = orng.Duplicate;
                            while (tmporng.Text.EndsWith(" "))
                            {
                                tmporng.SetRange(tmporng.Start, tmporng.End - 1);
                            }
                            while (tmporng.Text.StartsWith(" "))
                            {
                                tmporng.SetRange(tmporng.Start + 1, tmporng.End);
                            }                            
                            if (IsValidCitRange(tmporng.Text))
                            {
                                while (tmporng.Text.EndsWith("\r") ||
                                       tmporng.Text.EndsWith("\f") ||
                                       tmporng.Text.EndsWith("\v") ||
                                       tmporng.Text.EndsWith(",") ||
                                       tmporng.Text.EndsWith(";") ||
                                       tmporng.Text.EndsWith("."))
                                {
                                    tmporng.SetRange(tmporng.Start, tmporng.End - 1);
                                }
                                bookID++;
                                fndrefs.Add(new XrefModel()
                                {
                                    XrefText = tmporng.Text,
                                    XrefType = strType,
                                    XrefBookmark = "XREF_" + bookID,
                                    XrefContext = ClsCommonUtils.GetContextfromRng(tmporng.Duplicate)
                                });
                                actdoc.Bookmarks.Add("XREF_" + bookID, tmporng.Duplicate);
                            }
                            prvsrng = orng.Duplicate;
                            orng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
                return fndrefs;
            }
            catch { return null; }
        }

        private static bool IsNotValidCitRange(string wtext, Word.Range wRng)
        {
            bool notvalid = false;
            //string ChemicalElms = "|||H|||He|||Li|||Be|||B|||C|||N|||O|||F|||Ne|||Na|||Mg|||Al|||Si|||P|||S|||Cl|||Ar|||K|||Ca|||Sc|||Ti|||V|||Cr|||Mn|||Fe|||Co|||Ni|||Cu|||Zn|||Ga|||Ge|||As|||Se|||Br|||Kr|||Rb|||Sr|||Y|||Zr|||Nb|||Mo|||Tc|||Ru|||Rh|||Pd|||Ag|||Cd|||In|||Sn|||Sb|||Te|||I|||Xe|||Cs|||Ba|||Lu|||Hf|||Ta|||W|||Re|||Os|||Ir|||Pt|||Au|||Hg|||Tl|||Pb|||Bi|||Po|||At|||Rn|||Fr|||Ra|||Lr|||Rf|||Db|||Sg|||Bh|||Hs|||Mt|||Ds|||Rg|||Cn|||Uut|||Uuq|||Uup|||Uuh|||Uus|||Uuo|||La|||Ce|||Pr|||Nd|||Pm|||Sm|||Eu|||Gd|||Tb|||Dy|||Ho|||Er|||Tm|||Yb|||Ac|||Th|||Pa|||U|||Np|||Pu|||Am|||Cm|||Bk|||Cf|||Es|||Fm|||Md|||No|||";
            string stmptxt = wtext;
            if (stmptxt.Trim().StartsWith("-")) { return true; }
            return notvalid;
        }

        private static bool IsValidCitRange(string swText)
        {
            string[] strpuncs = new string[] { ",", " ", ";", "(", ")", "[", "]", "\u2013", "-" };
            string tmptext = swText;
            foreach (string s in strpuncs)
            {
                tmptext = tmptext.Replace(s, "");
            }
            if (int.TryParse(tmptext, out int n)) { return true; }
            return false;
        }
    }
}
