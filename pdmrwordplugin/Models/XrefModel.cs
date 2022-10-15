using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{
    public class XrefModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }
        public string XrefType { get; set; }
        public string XrefText { get; set; }
        public string XrefBookmark { get; set; }
        public string XrefContext { get; set; }

        private bool _XrefCheckbox;
        public bool XrefCheckbox
        {
            get { return _XrefCheckbox; }
            set
            {
                _XrefCheckbox = value;
                RaisePropertyChanged("XrefCheckbox");
            }
        }

        private bool _XrefSelected;
        public bool XrefSelected
        {
            get { return _XrefSelected; }
            set
            {
                _XrefSelected = value;
                RaisePropertyChanged("XrefSelected");
            }
        }
    }
}
