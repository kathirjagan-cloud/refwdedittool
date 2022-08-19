using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{
    public class Author
    {
        public string given { get; set; }
        public string family { get; set; }
    }

    public class ReferenceModel
    {
        public string Reftext { get; set; }
        public string Refbookmark { get; set; }
        public long PIndex { get; set; }
        public string RefJSON { get; set; }
        public string ReftextHtml { get; set; }
        public string RefStrucText { get; set; }
        public string RefCompText { get; set; }
        public List<Author> Authors { get; set; }
        public string title { get; set; }
        public string volume { get; set; }
        public string containertitle { get; set; }
        public string containertitleabbrv { get; set; }
        public string date { get; set; }
        public string issue { get; set; }
        public string pages { get; set; }
    }
}
