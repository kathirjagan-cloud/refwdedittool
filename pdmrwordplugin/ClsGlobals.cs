using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin
{
    public static class ClsGlobals
    {
        public const string REF_PARSER_EXE = "AnystyleParser.exe";
        public const string REF_PARSER_PATH = @"F:\Parser\";
        public const string REF_PARA_STYLE = "PReference";
        public const string PROJ_TITLE = "PDMR Word Plug-in v1.0";
        public static string APP_PATH = @"%appdata%\pdmrplugin\";
        public const string REF_STYLES_CONFIG = "ReferenceStyles.xml";
        public const string REF_BOOK_NAME = "REFMain";
        public const string XREF_SUP_STYLE_NAME = "pdmrspxref";
        public const string XREF_ONLINE_STYLE_NAME = "pdmronxref";
        public static Models.referencestyles gReferencestyles { set; get; }
    }
}
