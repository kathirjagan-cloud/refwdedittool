using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Utilities
{
    public static class ClsRefPub
    {
        public static Task IParseReferencebyExe(List<ReferencePostions> reflists)
        {
            return Task.Run(() => ParseReferencebyExe(reflists));
        }
        public static void ParseReferencebyExe(List<ReferencePostions> processreflists)
        {

        }
        public static void ParseReferencebyOnlinePub()
        {

        }
        public static void ParseReferencebyOnlineXref()
        {

        }
    }
}
