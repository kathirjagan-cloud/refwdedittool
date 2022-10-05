using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public XrefViewModel()
        {
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
