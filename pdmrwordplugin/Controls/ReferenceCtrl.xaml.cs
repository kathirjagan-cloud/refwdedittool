using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace pdmrwordplugin
{
    /// <summary>
    /// Interaction logic for ReferenceCtrl.xaml
    /// </summary>
    public partial class ReferenceCtrl : UserControl
    {
        public ReferenceCtrl()
        {
            InitializeComponent();
            cmbStyles.Items.Add("Test");
            cmbStyles.Items.Add("Mani good of having example right");
            List<ReferenceModel> references = Utilities.ClsRefPub.GetReferencesFromDoc();
            this.DataContext= new ViewModels.RefParserModel(references);
        }
    }
}
