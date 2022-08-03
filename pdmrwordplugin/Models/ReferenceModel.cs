using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{
    public class ReferenceModel
    {
        public string Reftext { get; set; }
        public string Refbookmark { get; set; }
        public long PIndex { get; set; }
        public string RefJSON { get; set; }
    }
}
