using DocumentFormat.OpenXml.Packaging;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using DOCXML = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System.IO;
using System.Management.Automation;
using System.Xml.Serialization;
using System.Windows.Markup;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using pdmrwordplugin.Functions;
using MaterialDesignColors.Recommended;
using System.Linq.Expressions;
using System.IO.Packaging;

namespace pdmrwordplugin.ViewModels
{
    public class RefParserModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }

        public RelayCommand NextReferenceCmd { set; get; }
        public RelayCommand SearchOnlineTermCmd { set; get; }
        public RelayCommand ApplyStyledRefs { get; set; }
       
        private bool _FirstRunRef;
        public bool FirstRunRef
        {
            get { return _FirstRunRef; }
            set
            {
                _FirstRunRef = value;
                RaisePropertyChanged("FirstRunRef");
            }
        }

        private bool _ShowRefProc;
        public bool ShowRefProc
        {
            get { return _ShowRefProc; }
            set
            {
                _ShowRefProc = value;
                RaisePropertyChanged("ShowRefProc");
            }
        }

        private bool _ShowActionButton;
        public bool ShowActionButton
        {
            get { return _ShowActionButton; }
            set
            {
                _ShowActionButton = value;
                RaisePropertyChanged("ShowActionButton");
            }
        }

        private ObservableCollection<ReferenceModel> _ProcessReferences;
        public ObservableCollection<ReferenceModel> ProcessReferences
        {
            get { return _ProcessReferences; }
            set
            {
                _ProcessReferences = value;
                RaisePropertyChanged("ProcessReferences");
            }
        }

        private string _SearchTextOnline;
        public string SearchTextOnline
        {
            get { return _SearchTextOnline; }
            set
            {
                _SearchTextOnline = value;
                RaisePropertyChanged("SearchTextOnline");
            }
        }

        private bool _ToShowSearch;
        public bool ToShowSearch
        {
            get { return _ToShowSearch; }
            set
            {
                _ToShowSearch = value;
                RaisePropertyChanged("ToShowSearch");
            }
        }

        private int _MainTabIndex;
        public int MainTabIndex
        {
            get { return _MainTabIndex; }
            set
            {
                _MainTabIndex = value;
                RaisePropertyChanged("MainTabIndex");
            }
        }

        private bool _showprogress;
        public bool Showprogress
        {
            get { return _showprogress; }
            set
            {
                _showprogress = value;
                RaisePropertyChanged("Showprogress");
            }
        }

        private int _SelTabIndex;
        public int SelTabIndex
        {
            get { return _SelTabIndex; }
            set
            {
                _SelTabIndex = value;
                RaisePropertyChanged("SelTabIndex");
                if (SelTabIndex == 1)
                {
                    BeginSearchTextinPubmed();
                }
            }
        }

        private ReferenceModel _SelReference;
        public ReferenceModel SelReference
        {
            get { return _SelReference; }
            set
            {
                _SelReference = value;
                if (value != null)
                {
                    value.ReftextHtml = GetFormatTextOpenXML(Globals.ThisAddIn.Application.Selection.Range.Duplicate);
                    if (SelTabIndex == 1) { BeginSearchTextinPubmed(); }
                }
                RaisePropertyChanged("SelReference");
            }
        }

        public void DoNextReference()
        {
            try
            {
                Word.Range nxtrng = Globals.ThisAddIn.Application.Selection.Range.Next(Word.WdUnits.wdParagraph, 1);
                if (nxtrng != null)
                {
                    nxtrng.Select();
                    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(Globals.ThisAddIn.Application.Selection.Range);
                    ReferenceModel selmodel = GetSelectedReference();
                    if (selmodel != null) { SelReference = selmodel; nxtrng.Select(); }
                    else { SelReference = null; }
                }
            }
            catch { }
        }

        private ReferenceModel GetSelectedReference()
        {
            try
            {
                ReferenceModel resreference = null;
                foreach (Word.Bookmark bk in Globals.ThisAddIn.Application.Selection.Range.Bookmarks)
                {
                    if (bk.Name.StartsWith("REF_"))
                    {
                        var selref = from item in ProcessReferences
                                     where item.Refbookmark.ToLower() == bk.Name.ToLower()
                                     select item;
                        if (selref != null && selref.Count() > 0)
                        { resreference = selref.FirstOrDefault(); break; }
                    }
                }
                return resreference;
            }
            catch { return null; }
        }

        public void BeginSearchTextinPubmed()
        {
            Showprogress = true;
            ShowActionButton = false;
            Utilities.clsRefSearchOnline.ISearchPubmed(SelReference).ContinueWith(t =>
            {
                ShowActionButton = true;
                Showprogress = false;
                if (!t.IsFaulted && t.Result != null)
                {
                    string taggedres= GetFormatTextPubmed(t.Result, SelReference.Reftext);
                    SelReference.ReftaggedText = taggedres;
                    if (!string.IsNullOrEmpty(taggedres)) { taggedres = taggedres.Replace("<title>", ""); taggedres = taggedres.Replace("</title>", ""); }
                    SelReference.RefStrucText = taggedres; //GetFormatTextPubmed(t.Result, SelReference.Reftext);
                    SelReference.RefCompText = GetCompareText(SelReference.Reftext, SelReference.RefStrucText);
                    RaisePropertyChanged("SelReference");
                }
            });
        }

        public string GetCompareText(string orgtext, string structext)
        {
            string flowdocstart = @"<FlowDocument xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">";
            string flowdocend = "</FlowDocument>";
            string puncts = "$$in$$was$$and$$";
            try
            {
                structext = structext.Replace(flowdocstart, "");
                structext = structext.Replace(flowdocend, "");
                structext = structext.Replace("<Paragraph>", "");
                structext = structext.Replace("</Paragraph>", "");
                structext = structext.Replace("<Bold>", "");
                structext = structext.Replace("</Bold>", "");
                structext = structext.Replace("<Italic>", "");
                structext = structext.Replace("</Italic>", "");
                string[] strarr1 = structext.Split(new char[] { ' ', '.', ':', ',', ';', '(', ')', '[', ']', '-', '\u2014' });
                var strarr = strarr1.Distinct().ToArray().OrderByDescending(x => x.Length).ToArray();
                for (int i = 0; i < strarr.Length; i++)
                {
                    string s = strarr[i];
                    if (s.Length > 3 && !puncts.Contains(s))
                    {
                        orgtext = orgtext.ReplaceFirst(s, "<Run Foreground=\"Green\">" + "$$" + i.ToString() + "$$" + "</Run>");
                    }
                    else if (int.TryParse(s, out _) && Regex.IsMatch(orgtext, "\b" + s))
                    {
                        orgtext = orgtext.ReplaceFirst(s, "<Run Foreground=\"Green\">" + "$$" + i.ToString() + "$$" + "</Run>");
                    }
                }
                for (int i = 0; i < strarr.Length; i++)
                {
                    orgtext = orgtext.ReplaceFirst("$$" + i.ToString() + "$$", strarr[i]);
                }
                return flowdocstart + "<Paragraph>" + orgtext + "</Paragraph>" + flowdocend;
            }
            catch { return ""; }
        }

        private static ReferenceModel GetPubmedObject(string parsexmlstr)
        {
            try
            {
                ReferenceModel classObject = null;
                if (parsexmlstr.Contains("<PubmedArticle>"))
                {
                    classObject = new ReferenceModel();
                    var xDoc = XDocument.Parse(parsexmlstr);
                    var article = xDoc.Root;
                    var authors = (from item in article.Descendants("Author")
                                   select new Author()
                                   {
                                       family = item.Element("LastName") != null ? item.Element("LastName").Value : "",
                                       given = item.Element("Initials") != null ? item.Element("Initials").Value : "",
                                       forename = item.Element("ForeName") != null ? item.Element("ForeName").Value : ""
                                   }).ToList();

                    if (authors != null) { classObject.Authors = new List<Author>(authors); }

                    var elem = xDoc.XPathSelectElement(".//Volume");
                    if (elem != null) { classObject.volume = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//Issue");
                    if (elem != null) { classObject.issue = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//PubDate/Year");
                    if (elem != null) { classObject.date = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//Journal/Title");
                    if (elem != null) { classObject.containertitle = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//MedlineJournalInfo/MedlineTA");
                    if (elem != null) { classObject.containertitleabbrv = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//Journal/ISOAbbreviation");
                    if (elem != null) { classObject.containertitleabbrv = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//MedlinePgn");
                    if (elem != null) { classObject.pages = elem.Value; }

                    elem = xDoc.XPathSelectElement(".//ArticleTitle");
                    if (elem != null) { classObject.title = elem.Value; }

                    return classObject;
                }
                else { return null; }
            }
            catch { return null; }
        }


        private static string GetFormatTextPubmed(string xmlstr, string refstext)
        {
            string flowdocstart = @"<FlowDocument xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">";
            string flowdocend = "</FlowDocument>";
            try
            {
                ReferenceModel pubmedobj = GetPubmedObject(xmlstr);
                if (pubmedobj == null) { return ""; }
                referencestylesStyle refstyle = ClsGlobals.gReferencestyles.style.FirstOrDefault();
                string refpattern = refstyle.pattern;
                string authorpattern = GetAuthorsFormat(refstyle, pubmedobj.Authors.Count);
                MatchCollection lastmatches = Regex.Matches(authorpattern, @"\[LastName\]");
                MatchCollection inimatches = Regex.Matches(authorpattern, @"\[Initials\]");                
                if (lastmatches != null && lastmatches.Count > 0)
                {
                    for (int i = 0; i < lastmatches.Count; i++)
                    {
                        authorpattern = StringExtensionMethods.ReplaceFirst(authorpattern, "[LastName]", pubmedobj.Authors[i].family);
                    }
                }
                if (inimatches != null && inimatches.Count > 0)
                {
                    for (int i = 0; i < inimatches.Count; i++)
                    {
                        authorpattern = StringExtensionMethods.ReplaceFirst(authorpattern, "[Initials]", pubmedobj.Authors[i].given);
                    }
                }

                refpattern = refpattern.Replace("[Author]", authorpattern);
                //replace author "[Author] [ArticleTitle]. [Journal] [Date];[Volume]([Issue]):[Page]"

                if (refstyle.journal.abbreviation)
                {
                    string tmpjtitle = "";
                    string strJtitle = pubmedobj.containertitleabbrv;
                    bool.TryParse(((System.Xml.XmlNode[])refstyle.journal.useperiod)[0].Value, out bool blnuseperiod);
                    if (!blnuseperiod) { strJtitle = strJtitle.Replace(".", ""); }
                    else
                    {
                        string orgjrnltxt = "**" + pubmedobj.containertitle.Replace(" ","**") + "**";
                        if (!strJtitle.Contains("."))
                        {
                            foreach (string s in strJtitle.Split(' '))
                            {
                                if (!orgjrnltxt.Contains("**" + s + "**")) { tmpjtitle = tmpjtitle + s + ". "; }
                                else { tmpjtitle = tmpjtitle + s + " "; }
                            }
                        }
                        if (!string.IsNullOrEmpty(tmpjtitle)) { strJtitle = tmpjtitle; }
                    }                    
                    refpattern = refpattern.Replace("[Journal]", GetFormattingbyStyle(false, refstyle.journal.italic, strJtitle));
                }
                else
                {                    
                    refpattern = refpattern.Replace("[Journal]", GetFormattingbyStyle(false, refstyle.journal.italic, pubmedobj.containertitle));
                }
                //replace journal title

                string tmpatitle = pubmedobj.title;
                if (tmpatitle.EndsWith(".")) { tmpatitle = tmpatitle.Substring(0, tmpatitle.Length - 1); }
                refpattern = refpattern.Replace("[ArticleTitle]", "<title>" + tmpatitle + "</title>");
                //replace article title
                refpattern = refpattern.Replace("[Date]", GetFormattingbyStyle(refstyle.date.bold, refstyle.date.italic, pubmedobj.date));
                //replace date
                refpattern = refpattern.Replace("[Volume]", GetFormattingbyStyle(refstyle.volume.bold, refstyle.volume.italic, pubmedobj.volume));
                //replace issue
                refpattern = refpattern.Replace("[Issue]", GetFormattingbyStyle(refstyle.issue.bold, refstyle.issue.italic, pubmedobj.issue));
                //replace volume
                refpattern = refpattern.Replace("[Page]", GetEllidedNumberbyStyle(refstyle.page.ellision, refstyle.page.separator, refstyle.page.omitchars, pubmedobj.pages));
                //replace pagination
                refpattern = refpattern.Replace("&", "&amp;");
                return flowdocstart + "<Paragraph>" + GetReferenenumber(refstext) + refpattern + "</Paragraph>" + flowdocend;
            }
            catch
            {
                return "";
            }
        }


        private static string GetReferenenumber(string strreftxt)
        {
            try
            {
                string result = "";
                string rgxpattern = "^([ \\t]+)?([\\(\\[])?(([0-9]+)([ .\\t]+)?)+([\\)\\]])?";
                if (Regex.IsMatch(strreftxt, rgxpattern))
                {
                    result = Regex.Match(strreftxt, rgxpattern).Value;
                    result = result.Replace(" ", "");
                    result = result.Replace("[", "");
                    result = result.Replace("]", "");
                    result = result.Replace("(", "");
                    result = result.Replace(")", "");
                    result = result.Replace(".", "");
                    result = result.Replace("\t", "");
                }
                if (!string.IsNullOrEmpty(result)) { result += ". "; }
                return result;
            }
            catch { return ""; }
        }

        private static string GetEllidedNumberbyStyle(bool toEllide, string eSepr, string removechar, string strpage)
        {
            try
            {
                string frstnum = ""; string sndnum = "";
                string tmppage = strpage;
                string result = "";
                tmppage = tmppage.Replace("\u2013", "-");
                tmppage = tmppage.Replace("\u2014", "-");
                if (tmppage.Contains("-"))
                {
                    string[] numspl = tmppage.Split('-');
                    frstnum = numspl[0]; sndnum = numspl[1];
                    if (!toEllide)
                    {
                        if (frstnum.Length > sndnum.Length)
                        {
                            sndnum = frstnum.Substring(0, frstnum.Length - sndnum.Length) + sndnum;
                        }
                        result = frstnum + eSepr + sndnum;
                    }
                    else
                    {
                        if (frstnum.Length > 1 && sndnum.Length > 1 && frstnum.Length == sndnum.Length)
                        {
                            int matched = 0;
                            for (int k = 0; k < frstnum.Length; k++)
                            {
                                matched++;
                                if (frstnum.Substring(k, 1) != sndnum.Substring(k, 1)) { break; }
                            }
                            if (frstnum.Length != sndnum.Substring(matched - 1, sndnum.Length - (matched - 1)).Length)
                            {
                                foreach (string sp in removechar.Split('|'))
                                {
                                    if (!string.IsNullOrEmpty(sp))
                                        frstnum = frstnum.Replace(sp, "");
                                }
                            }
                            result = frstnum + eSepr + sndnum.Substring(matched - 1, sndnum.Length - (matched - 1));
                        }
                        else
                        {
                            result = frstnum + eSepr + sndnum;
                        }
                    }
                }
                if (string.IsNullOrEmpty(result)) { result = strpage; }
                return result;
            }
            catch { return strpage; }
        }

        private static string GetFormattingbyStyle(bool blnbold, bool blnitalic, string val)
        {
            try
            {
                if(blnbold && blnitalic) { val = "<Bold><Italic>" + val + "<Italic></Bold>"; }
                if (blnbold) { val = "<Bold>" + val + "</Bold>"; }
                if (blnitalic) { val = "<Italic>" + val + "</Italic>"; }
                return val;
            }
            catch { return val; }
        }

        private static string GetAuthorsFormat(referencestylesStyle arefstyle, int authorscount)
        {
            string authstr = arefstyle.authorpattern;
            try
            {
                int maxcount = 0;
                int count = 0;
                string authorendsp = arefstyle.separators.end;
                if (!string.IsNullOrEmpty(arefstyle.separators.maxcount))
                {
                    maxcount = int.Parse(arefstyle.separators.maxcount);
                }
                if (!string.IsNullOrEmpty(arefstyle.separators.count))
                {
                    count = int.Parse(arefstyle.separators.count);
                }
                if (maxcount > 0 && authorscount < maxcount)
                {
                    count = authorscount;
                }
                else if (maxcount > 0 && authorscount > maxcount && !string.IsNullOrEmpty(arefstyle.separators.etal))
                {
                    authorendsp = arefstyle.separators.etal;
                }
                if (count == 2 && !string.IsNullOrEmpty(arefstyle.separators.twoauthor))
                {
                    authstr = authstr + arefstyle.separators.twoauthor + authstr + arefstyle.separators.end;
                }
                else if (count > 1)
                {
                    string strtmp = "";
                    for (int i = 1; i <= count; i++)
                    {
                        if (string.IsNullOrEmpty(strtmp)) { strtmp = authstr; }
                        else { strtmp += arefstyle.separators.author + authstr; }
                    }
                    authstr = strtmp + authorendsp;
                }
                return authstr;
            }
            catch
            {
                return authstr;
            }
        }


        private static string GetFormatTextOpenXML(Word.Range orange)
        {
            string flowdocstart = @"<FlowDocument xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">";
            string flowdocend = "</FlowDocument>";
            try
            {
                string strXaml = "";
                string FlOpenXml = System.IO.Path.GetTempPath() + "Openflat.xml";
                orange.ExportFragment(FlOpenXml, Word.WdSaveFormat.wdFormatXMLDocument);
                WordprocessingDocument document = WordprocessingDocument.Open(FlOpenXml, false);
                Body docbody = document.MainDocumentPart.Document.Body;
                foreach (DOCXML.Wordprocessing.Paragraph para in docbody.OfType<DOCXML.Wordprocessing.Paragraph>())
                {
                    foreach (var run in para.Descendants())
                    {
                        string runtext = "";
                        if (run.XName == W.r)
                        {
                            if (!run.OuterXml.Contains("<w:delText"))
                            {
                                foreach (var text in run.Descendants<Text>())
                                {
                                    runtext += text.Text;
                                }
                            }
                        }
                        if (run.OuterXml.Contains("<w:b />") || run.OuterXml.Contains("<w:b/>"))
                        {
                            strXaml += "<Bold>" + runtext + "</Bold>";
                        }
                        if (run.OuterXml.Contains("<w:i />") || run.OuterXml.Contains("<w:i/>"))
                        {
                            strXaml += "<Italic>" + runtext + "</Italic>";
                        }
                        else
                        {
                            strXaml += runtext;
                        }
                    }
                }
                strXaml = strXaml.Replace(" <", "\u2002<");
                strXaml = strXaml.Replace("&", "&amp;");
                document.Close();
                return flowdocstart + "<Paragraph>" + strXaml + "</Paragraph>" + flowdocend;
            }
            catch
            {
                return flowdocstart + "<Paragraph>" + orange.Text + "</Paragraph>" + flowdocend;
            }
        }

        public void ToSearchOnline()
        {
            ToShowSearch = true;
            MainTabIndex = 1;
            if (SelReference != null)
                SearchTextOnline = "https://www.google.com/search?q=" + SelReference.Reftext;
        }

        public void FormatReferenceinSel()
        {
            try
            {
                if (SelReference == null) return;
                if (SelReference.ReftextHtml != null)
                {
                    string orngBK = SelReference.Refbookmark;
                    Word.Range orng = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                    if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(orngBK))
                    {
                        orng = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[orngBK].Range.Duplicate;
                    }
                    string scomptxt = ClsCommonUtils.GetCleanPubmedTxt(SelReference.RefStrucText);                    
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    #region Using Normal way
                    //ClsCommonUtils.TRACK_OFF();
                    //orng.Text = ClsCommonUtils.ShowDifferences(SelReference.Reftext, scomptxt, true);
                    //ClsCommonUtils.TRACK_ON();
                    //Globals.ThisAddIn.Application.ActiveWindow.View.RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal;
                    //ClsCommonUtils.InsertwithTrack(orng.Duplicate);
                    //ClsCommonUtils.TRACK_OFF();
                    //ClsCommonUtils.ClearTagsinSelection(orng.Duplicate);
                    //orng.Select();
                    #endregion
                    
                    ClsCommonUtils.TRACK_OFF();
                    Word.Document cmpdoc = ClsCommonUtils.GetCompareRangeDoc(SelReference.Reftext, scomptxt, SelReference.ReftaggedText);
                    if (cmpdoc != null)
                    {
                        Word.Range cmprng = cmpdoc.Paragraphs[1].Range.Duplicate;
                        cmprng.SetRange(cmprng.Start, cmprng.End - 1);
                        orng.FormattedText = cmprng.Duplicate;
                        cmpdoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    }                    
                    orng.Select();                    
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
            }
            catch
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }        

        public void DoProcessReferences()
        {
            try
            {
                Showprogress = true;
                ShowActionButton = false;
                Utilities.ClsRefPub.IParseReferencebyExe(docreferences).ContinueWith(t =>
                {
                    Showprogress = false;
                    ShowActionButton = true;
                    if (!t.IsFaulted && t.Result != null)
                    {
                        MainTabIndex = 1;
                        FirstRunRef = false;
                        ShowRefProc = true;
                        ProcessReferences = new ObservableCollection<ReferenceModel>(t.Result);
                        Globals.ThisAddIn.Application.Selection.Paragraphs.First.Range.Select();
                        ReferenceModel defaultmodl = ProcessReferences.FirstOrDefault();
                        if (defaultmodl != null)
                        {
                            if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(defaultmodl.Refbookmark))
                            {                                
                                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[defaultmodl.Refbookmark].Range.Select();
                                Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(Globals.ThisAddIn.Application.Selection.Range);
                            }
                        }
                        SelReference = defaultmodl;
                        if (defaultmodl != null)
                        {
                            if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(defaultmodl.Refbookmark))
                            {
                                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[defaultmodl.Refbookmark].Range.Select();
                                Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(Globals.ThisAddIn.Application.Selection.Range);
                            }
                        }
                    }
                });
            }
            catch { }
        }

        private List<ReferenceModel> docreferences { set; get; }

        public RelayCommand StartRefProcess { get; set; }

        #region Initialize

        #endregion

        public RefParserModel(List<ReferenceModel> _docreferences)
        {
            NextReferenceCmd = new RelayCommand(m => DoNextReference());
            SearchOnlineTermCmd = new RelayCommand(m => ToSearchOnline());
            ApplyStyledRefs = new RelayCommand(m => FormatReferenceinSel());
            ProcessReferences = new ObservableCollection<ReferenceModel>();
            StartRefProcess = new RelayCommand(m => DoProcessReferences());
            if (_docreferences == null) { _docreferences = new List<ReferenceModel>(); }
            docreferences = new List<ReferenceModel>(_docreferences);
            FirstRunRef = true;
            ShowRefProc = false;
            MainTabIndex = 0;
        }
    }
}
