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

namespace pdmrwordplugin.ViewModels
{
    public class RefParserModel: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
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
                    Showprogress = true;
                    Utilities.clsRefSearchOnline.ISearchPubmed(SelReference).ContinueWith(t =>
                    {
                        Showprogress = false;
                        if(!t.IsFaulted && t.Result!=null)
                        {
                            string s = t.Result;
                        }
                    });
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
                    Globals.ThisAddIn.Application.Selection.Paragraphs.First.Range.Select();
                    value.ReftextHtml = GetFormatTextOpenXML(Globals.ThisAddIn.Application.Selection.Range.Duplicate);
                }
                RaisePropertyChanged("SelReference");
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
                document.Close();
                return flowdocstart + "<Paragraph>" + strXaml + "</Paragraph>" + flowdocend;
            }
            catch
            {
                return flowdocstart + "<Paragraph>" + orange.Text + "</Paragraph>" + flowdocend;
            }
        }

        private static string GetFormatText(Word.Range orange)
        {
            string flowdocstart = @"<FlowDocument xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"">";
            string flowdocend = "</FlowDocument>";
            try
            {
                string strXaml = "";
                foreach (Word.Range orng in orange.Characters)
                {
                    bool blnItalic = false;
                    bool blnBold = false;
                    if (orng.Font.Bold != 0)
                    {
                        blnBold = true;
                        strXaml += "<Bold>" + orng.Text + "</Bold>";
                    }
                    if (orng.Font.Italic != 0)
                    {
                        blnItalic = true;
                        strXaml += "<Italic>" + orng.Text + "</Italic>";
                    }
                    if (!blnBold && !blnItalic)
                        strXaml += orng.Text;
                }
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
            ProcessReferences = new ObservableCollection<ReferenceModel>();
            Showprogress = true;
            Utilities.ClsRefPub.IParseReferencebyExe(docreferences).ContinueWith(t =>
            {
                Showprogress = false;
                if (!t.IsFaulted && t.Result != null)
                {
                    ProcessReferences = new ObservableCollection<ReferenceModel>(t.Result);
                    SelReference = ProcessReferences.FirstOrDefault();
                    //SelReference.ReftextHtml = @"<FlowDocument xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""><Paragraph><Bold>Hello World!</Bold></Paragraph></FlowDocument>";
                }
            });
        }
    }
}
