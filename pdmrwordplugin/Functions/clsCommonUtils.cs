using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.Functions
{
    public static class ClsCommonUtils
    {
        // Get the reference details like text, bookmark //
        public static List<ReferencePostions> GetReferenceDetails(Word.Document document, Word.Range Selectedrange = null)
        {
            try
            {
                long paraindex = 0;
                int bookindex = 0;
                List<ReferencePostions> objpos = new List<ReferencePostions>();
                Word.Range processrange = document.Range().Duplicate;
                if (Selectedrange != null) { processrange = Selectedrange; }
                foreach (Word.Paragraph paragraph in processrange.Paragraphs)
                {
                    paraindex++;
                    string parastyle = GetRangeStyleName(paragraph.Range);
                    if (!string.IsNullOrEmpty(parastyle) &&
                        parastyle == ClsGlobals.REF_PARA_STYLE &&
                        !string.IsNullOrEmpty(paragraph.Range.Text))
                    {
                        bookindex++;
                        objpos.Add(new ReferencePostions()
                        {
                            Reftext = paragraph.Range.Text,
                            Refbookmark = "_REF_" + String.Format("{0:D4}", bookindex),
                            PIndex = paraindex
                        });
                    }
                }
                processrange = null;
                return objpos;
            }
            catch
            {
                return null;
            }
        }

        public static string GetRangeStyleName(Word.Range range)
        {
            // Get the style of the paragraph
            try
            {
                return (string)range.get_Style();
            }
            catch { return null; }
        }

    }
}
