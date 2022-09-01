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

namespace pdmrwordplugin.ViewModels
{
    public class RefParserModel: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }

        public RelayCommand NextReferenceCmd { set; get; }

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
                    ReferenceModel selmodel = GetSelectedReference();
                    if (selmodel != null) { SelReference = selmodel; }
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
            Utilities.clsRefSearchOnline.ISearchPubmed(SelReference).ContinueWith(t =>
            {
                Showprogress = false;
                if (!t.IsFaulted && t.Result != null)
                {
                    SelReference.RefStrucText = GetFormatTextPubmed(t.Result);
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
                    
                    elem = xDoc.XPathSelectElement(".//MedlinePgn");
                    if (elem != null) { classObject.pages= elem.Value; }
                    
                    elem = xDoc.XPathSelectElement(".//ArticleTitle");
                    if (elem != null) { classObject.title = elem.Value; }

                    return classObject;
                }
                else { return null; }
            }
            catch { return null; }
        }
       

        private static string GetFormatTextPubmed(string xmlstr)
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
                    if (refstyle.journal.italic) { refpattern = refpattern.Replace("[Journal]", "<Italic>" + pubmedobj.containertitleabbrv + "</Italic>"); }
                    else { refpattern = refpattern.Replace("[Journal]", pubmedobj.containertitleabbrv); }
                }
                else
                {
                    if (refstyle.journal.italic) { refpattern = refpattern.Replace("[Journal]", "<Italic>" + pubmedobj.containertitle + "</Italic>"); }
                    else { refpattern = refpattern.Replace("[Journal]", pubmedobj.containertitle); }
                }
                //replace journal title
                refpattern = refpattern.Replace("[ArticleTitle]", pubmedobj.title);
                //replace article title
                refpattern = refpattern.Replace("[Date]", pubmedobj.date);
                //replace date
                refpattern = refpattern.Replace("[Volume]", pubmedobj.volume);
                //replace issue
                refpattern = refpattern.Replace("[Issue]", pubmedobj.issue);
                //replace volume
                refpattern = refpattern.Replace("[Page]", pubmedobj.pages);
                //replace pagination
                refpattern = refpattern.Replace("&", "&amp;");
                return flowdocstart + "<Paragraph>" + refpattern + "</Paragraph>" + flowdocend;
            }
            catch
            {
                return "";
            }
        }

        private static string GetAuthorsFormat(referencestylesStyle arefstyle, int authorscount)
        {
            string authstr = arefstyle.authorpattern;
            try
            {
                int maxcount = 0;
                int count = 0;
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
                    arefstyle.separators.end = arefstyle.separators.etal;
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
                    authstr = strtmp + arefstyle.separators.end;
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

        #region Initialize

        #endregion

        public RefParserModel(List<ReferenceModel> docreferences)
        {
            NextReferenceCmd = new RelayCommand(m => DoNextReference());
            ProcessReferences = new ObservableCollection<ReferenceModel>();
            Showprogress = true;
            Utilities.ClsRefPub.IParseReferencebyExe(docreferences).ContinueWith(t =>
            {
                Showprogress = false;
                if (!t.IsFaulted && t.Result != null)
                {
                    ProcessReferences = new ObservableCollection<ReferenceModel>(t.Result);
                    Globals.ThisAddIn.Application.Selection.Paragraphs.First.Range.Select();
                    SelReference = ProcessReferences.FirstOrDefault();
                }
            });
        }
    }
}
