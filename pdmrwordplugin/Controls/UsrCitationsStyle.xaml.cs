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

namespace pdmrwordplugin.Controls
{
    /// <summary>
    /// Interaction logic for UsrCitationsStyle.xaml
    /// </summary>
    public partial class UsrCitationsStyle : UserControl
    {
        public UsrCitationsStyle()
        {
            InitializeComponent();
            cmbPositions.Items.Add("Online");
            cmbPositions.Items.Add("Superscript");
            cmbSeparators.Items.Add("Space");
            cmbSeparators.Items.Add("Comma");
            cmbSeparators.Items.Add("Comma+Space");
            cmbSeparators.Items.Add("Semicolon");
            cmbSeparators.Items.Add("Semicolon+Space");
            cmbRangeSeparators.Items.Add("Never");
            cmbRangeSeparators.Items.Add("en-dash");
            cmbhTimes.Items.Add("3");
            cmbhTimes.Items.Add("4");
            cmbhTimes.Items.Add("5");
            cmbhTimes.Items.Add("6");
            cmbhTimes.Items.Add("7");
            cmbRangeSeparators.Items.Add("hyphen");
            cmbresides.Items.Add("after punctuation");
            cmbresides.Items.Add("before punctuation");
            cmbBrackets.Items.Add("none");
            cmbBrackets.Items.Add("parentheses ()");
            cmbBrackets.Items.Add("brackets []");
        }
    }
}
