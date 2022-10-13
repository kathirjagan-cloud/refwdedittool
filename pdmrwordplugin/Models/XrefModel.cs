using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{
    public class XrefModel
    {
        public string XrefType { get; set; }
        public string XrefText { get; set; }
        public string XrefBookmark { get; set; }
        public string XrefContext { get; set; }
        public bool XrefSelected { get; set; }
    }
}
