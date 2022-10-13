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

        public void SelectAllXref()
        {
            if (SuperXrefs != null)
            {
                var list = SuperXrefs.ToList();
                list.All(x => { x.XrefSelected = true; return true; });
                SuperXrefs = new ObservableCollection<XrefModel>(list);                
            }
        }

        public void UnSelectAllXref()
        {
            if (SuperXrefs != null)
            {
                var list = SuperXrefs.ToList();
                list.All(x => { x.XrefSelected = false; return true; });
                SuperXrefs = new ObservableCollection<XrefModel>(list);
            }
        }

        public void ApplyXrefMark(object param)
        {
            if(param is XrefModel xref)
            {
                if(Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(xref.XrefBookmark))
                {
                    Word.Range orange = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[xref.XrefBookmark].Range.Duplicate;
                    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(orange);                    
                }
            }
        }

        public XrefViewModel()
        {
            CmdToselectAll = new RelayCommand(m => SelectAllXref());
            CmdToUnselectAll = new RelayCommand(m => UnSelectAllXref());
            CmdMarkSelection = new RelayCommand(m => ApplyXrefMark(m));
            Showprogress = true;
            Utilities.clsVanXref.ReadCitationsfromDoc().ContinueWith(t =>
            {
                Showprogress = false;
                if (!t.IsFaulted && t.Result != null)
                    SuperXrefs = new ObservableCollection<XrefModel>(t.Result);
            });
        }
    }
}
