using pdmrwordplugin.Functions;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.ViewModels
{
    public class XrefViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }

        private string _progresstext;
        public string Progresstext
        {
            get { return _progresstext; }
            set
            {
                _progresstext = value;
                RaisePropertyChanged("Progresstext");
            }
        }

        private ObservableCollection<XrefModel> _SuperXrefs;
        public ObservableCollection<XrefModel> SuperXrefs
        {
            get { return _SuperXrefs; }
            set
            {
                _SuperXrefs = value;
                RaisePropertyChanged("SuperXrefs");
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

        public RelayCommand CmdToselectAll { get; set; }
        public RelayCommand CmdToUnselectAll { get; set; }
        public RelayCommand CmdMarkSelection { get; set; }
        public RelayCommand CmdMarkAllSelection { get; set; }

        public void SelectAllXref()
        {
            if (SuperXrefs != null)
            {
                var list = SuperXrefs.ToList();
                list.All(x => { x.XrefCheckbox = true; return true; });
                SuperXrefs = new ObservableCollection<XrefModel>(list);                
            }
        }

        public void UnSelectAllXref()
        {
            if (SuperXrefs != null)
            {
                var list = SuperXrefs.ToList();
                list.All(x => { x.XrefCheckbox = false; return true; });
                SuperXrefs = new ObservableCollection<XrefModel>(list);
            }
        }


        private void ApplySelectedorAll(XrefModel xrefmod)
        {
            try
            {
                Word.Range orange = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[xrefmod.XrefBookmark].Range.Duplicate;
                Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(orange);
                if (!xrefmod.XrefSelected)
                {
                    ClsCommonUtils.SetStyelinRange(orange, ClsGlobals.XREF_SUP_STYLE_NAME);
                }
                else { ClsCommonUtils.SetStyelinRange(orange, ""); }                
            }
            catch { }
        }


        public void ApplyXrefMarkAll()
        {
            if(SuperXrefs!=null)
            {
                foreach(XrefModel xmod in SuperXrefs)
                {
                    if (xmod.XrefCheckbox)
                    {
                        ApplySelectedorAll(xmod);
                        xmod.XrefSelected = !xmod.XrefSelected;                        
                    }
                    xmod.XrefCheckbox = false;
                }
            }
        }

        public void ApplyXrefMark(object param)
        {
            if(param is XrefModel xref)
            {
                if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(xref.XrefBookmark))
                {
                    ApplySelectedorAll(xref);
                    xref.XrefSelected = !xref.XrefSelected;                    
                }
            }
        }        

        public XrefViewModel()
        {
            CmdToselectAll = new RelayCommand(m => SelectAllXref());
            CmdToUnselectAll = new RelayCommand(m => UnSelectAllXref());
            CmdMarkSelection = new RelayCommand(m => ApplyXrefMark(m));
            CmdMarkAllSelection = new RelayCommand(m => ApplyXrefMarkAll());
            Showprogress = true;
            IProgress<string> xprogress = new Progress<string>(s =>
            {
                Progresstext = s;
            });
            Utilities.clsVanXref.ReadCitationsfromDoc(xprogress).ContinueWith(t =>
            {
                Showprogress = false;                
                if (!t.IsFaulted && t.Result != null)
                    SuperXrefs = new ObservableCollection<XrefModel>(t.Result);
            });
        }
    }
}
