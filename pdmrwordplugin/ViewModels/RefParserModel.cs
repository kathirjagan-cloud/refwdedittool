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
                    ProcessReferences = new ObservableCollection<ReferenceModel>(t.Result);
            });
        }
    }
}
