using CiteProc.v10;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace pdmrwordplugin.Functions
{
    public static class ClsCommonUtils
    {

        public static T DeserializeXMLFileToObject<T>(string XmlFilename)
        {
            T returnObject = default(T);
            if (string.IsNullOrEmpty(XmlFilename)) return default(T);

            try
            {
                StreamReader xmlStream = new StreamReader(XmlFilename);
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                returnObject = (T)serializer.Deserialize(xmlStream);
            }
            catch
            {

            }
            return returnObject;
        }

        public static double CalculateSimilarity(string source, string target)
        {
            if ((source == null) || (target == null)) return 0.0;
            if ((source.Length == 0) || (target.Length == 0)) return 0.0;
            if (source.Trim().ToLower() == target.Trim().ToLower()) return 1.0;
            source = source.Replace(".", "").ToLower().Replace(" ", "");
            target = target.Replace(".", "").ToLower().Replace(" ", "");
            int stepsToSame = ComputeLevenshteinDistance(source, target);
            return (1.0 - ((double)stepsToSame / (double)Math.Max(source.Length, target.Length)));
        }

        public static int ComputeLevenshteinDistance(string source, string target)
        {
            if ((source == null) || (target == null)) return 0;
            if ((source.Length == 0) || (target.Length == 0)) return 0;
            if (source == target) return source.Length;

            int sourceWordCount = source.Length;
            int targetWordCount = target.Length;

            // Step 1
            if (sourceWordCount == 0)
                return targetWordCount;

            if (targetWordCount == 0)
                return sourceWordCount;

            int[,] distance = new int[sourceWordCount + 1, targetWordCount + 1];

            // Step 2
            for (int i = 0; i <= sourceWordCount; distance[i, 0] = i++) ;
            for (int j = 0; j <= targetWordCount; distance[0, j] = j++) ;

            for (int i = 1; i <= sourceWordCount; i++)
            {
                for (int j = 1; j <= targetWordCount; j++)
                {
                    // Step 3
                    int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;

                    // Step 4
                    distance[i, j] = Math.Min(Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1), distance[i - 1, j - 1] + cost);
                }
            }

            return distance[sourceWordCount, targetWordCount];
        }

        public static void TRACK_ON()
        {
            try
            {                
                Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = true;
            }
            catch { }
        }

        public static void TRACK_OFF()
        {
            try
            {
                Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = false;
            }
            catch { }
        }


        public static void SetReferenceRangebyBook()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveDocument == null) { return; }
                Word.Range orng = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                if (orng.Paragraphs.Count <= 5)
                {
                    MessageBox.Show(ClsMessages.REF_MESSAGE_4, ClsGlobals.PROJ_TITLE);
                }
                bool revstateindoc = Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions;
                ClsCommonUtils.TRACK_OFF();
                orng.ListFormat.ConvertNumbersToText();
                Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = revstateindoc;
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(ClsGlobals.REF_BOOK_NAME, orng.Duplicate);
            }
            catch { MessageBox.Show(ClsMessages.REF_MESSAGE_3, ClsGlobals.PROJ_TITLE); }
        }

        public static Word.Range GetReferenceRangebyBook()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveDocument == null) { return null; }
                Word.Document odoc = Globals.ThisAddIn.Application.ActiveDocument;
                if(odoc.Bookmarks.Exists(ClsGlobals.REF_BOOK_NAME))
                {
                    return odoc.Bookmarks[ClsGlobals.REF_BOOK_NAME].Range.Duplicate;
                }
                else
                {
                    MessageBox.Show(ClsMessages.REF_MESSAGE_5, ClsGlobals.PROJ_TITLE);
                }
                return null;
            }
            catch { return null; }
        }

        public static bool IsXrefStylePresent(string wstylname)
        {
            try
            {
                string stmp = Globals.ThisAddIn.Application.ActiveDocument.Styles[wstylname].Description;
                return true;
            }
            catch { return false; }
        }

        public static bool CreateXrefStyles()
        {
            try
            {
                bool isSuperscript = true;
                string stylename = ClsGlobals.XREF_SUP_STYLE_NAME;
                if (!IsXrefStylePresent(stylename))
                {
                    if (!CreateStylebyName(stylename, isSuperscript))
                        return false;
                }
                isSuperscript = false;
                stylename = ClsGlobals.XREF_ONLINE_STYLE_NAME;
                if (!IsXrefStylePresent(stylename))
                {
                    if (!CreateStylebyName(stylename, isSuperscript))
                        return false;
                }
                return true;
            }
            catch { return false; }
        }

        private static bool CreateStylebyName(string styletitle, bool isfmtSuper)
        {
            try
            {
                Word.Style ostyl = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add(styletitle, Word.WdStyleType.wdStyleTypeCharacter);
                if (isfmtSuper)
                    ostyl.Font.Superscript = 1;
                ostyl.Font.Bold = 0;
                foreach (Word.Border bdr in ostyl.Font.Borders)
                {
                    bdr.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    bdr.Color = Word.WdColor.wdColorSeaGreen;
                    bdr.LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                }
                return true;
            }
            catch { return false; }
        }

        public static string GetContextfromRng(Word.Range objrng)
        {
            try
            {
                string stext = "";
                Word.Range opararng = objrng.Paragraphs.First.Range.Duplicate;
                Word.Range moverng = objrng.Duplicate;
                moverng.Move(Word.WdUnits.wdCharacter, -75);
                if (!moverng.InRange(opararng))
                {
                    moverng = opararng.Duplicate;
                    moverng.SetRange(moverng.Start, objrng.End);
                }
                else
                    moverng.SetRange(moverng.Start, objrng.End);
                stext = moverng.Text;
                return stext;
            }
            catch { return objrng.Text; }
        }

        // Get the reference details like text, bookmark //
        public static List<ReferenceModel> GetReferenceDetails(Word.Document document, Word.Range Selectedrange = null)
        {
            try
            {
                long paraindex = 0;
                int bookindex = 0;
                List<ReferenceModel> objrefs = new List<ReferenceModel>();
                Word.Range processrange = document.Range().Duplicate;
                if (Selectedrange != null) { processrange = Selectedrange; }
                RemoveRefBookMarks(document);
                foreach (Word.Paragraph paragraph in processrange.Paragraphs)
                {
                    paraindex++;
                    string parastyle = GetRangeStyleName(paragraph.Range);
                    if (!string.IsNullOrEmpty(paragraph.Range.Text))
                    {
                        bookindex++;
                        Word.Range bkrng = paragraph.Range.Duplicate;
                        bkrng.SetRange(bkrng.Start, bkrng.End - 1);
                        document.Bookmarks.Add("REF_" + String.Format("{0:D4}", bookindex), bkrng);
                        objrefs.Add(new ReferenceModel()
                        {
                            Reftext = paragraph.Range.Text,
                            Refbookmark = "REF_" + String.Format("{0:D4}", bookindex),
                            PIndex = paraindex,
                            ReftextHtml = ""
                        });
                    }
                }
                processrange = null;
                return objrefs;
            }
            catch
            {
                return null;
            }
        }

        private static void RemoveRefBookMarks(Word.Document bkdoc)
        {
            try
            {
                for(int i= bkdoc.Bookmarks.Count; i >= 1; i--)
                {
                    bkdoc.Bookmarks[i].Delete();
                }
            }
            catch { }
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

        public static void SetStyelinRange(Word.Range wrng, string wstyle)
        {
            int bold = wrng.Font.Bold;
            int italic = wrng.Font.Italic;
            int superscript = wrng.Font.Superscript;
            if (!string.IsNullOrEmpty(wstyle))
                wrng.set_Style(wstyle);
            else
                wrng.Font.Reset();
            wrng.Font.Bold = bold;
            wrng.Font.Italic = italic;
            wrng.Font.Superscript = superscript;
        }

        public static string GetEllidedNumbers(string stext)
        {
            try
            {
                var result = string.Empty;
                if (stext.Contains(","))
                {
                    List<int> numberslist = new List<int>();
                    foreach (string s in stext.Split(','))
                    {
                        if (int.TryParse(s, out int x))
                        {
                            numberslist.Add(x);
                        }
                    }
                    if (numberslist.Count > 0)
                    {
                        int[] numbers = numberslist.ToArray();
                        int temp = numbers[0], start, end;
                        for (int i = 0; i < numbers.Length; i++)
                        {
                            start = temp;
                            if (i < numbers.Length - 1)                                
                                if (numbers[i] + 1 == numbers[i + 1])
                                    continue;                                
                                else
                                    temp = numbers[i + 1];
                            end = numbers[i];
                            if (start == end)
                                result += "," + start.ToString();
                            else
                                result += "," + start.ToString() + " - " + end.ToString();
                        }
                    }
                    if (result.StartsWith(",")) { result=result.Substring(1, result.Length - 1); }
                    return result;
                }
                else { return stext; }
            }
            catch { return stext; }
        }

        public static string GetReferenecesbyArea()
        {
            try
            {
                List<int> refnumbers = new List<int>();
                string rfpattern = @"^([ \\t]+)?([\\(\\[])?(([0-9]+)([ .\\t]+)?)+([\\(\\]])?";
                Regex regex = new Regex(rfpattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                Word.Range refrange = GetReferenceRangebyBook();                
                if (refrange != null)
                {
                    refrange.ListFormat.ConvertNumbersToText();
                    foreach (Word.Paragraph opara in refrange.Paragraphs)
                    {
                        string wholereftext = opara.Range.Text;
                        if (regex.IsMatch(wholereftext))
                        {
                            string rfnumber = regex.Matches(wholereftext)[0].Value;
                            rfnumber = rfnumber.Replace("[", "");
                            rfnumber = rfnumber.Replace("]", "");
                            rfnumber = rfnumber.Replace("(", "");
                            rfnumber = rfnumber.Replace(")", "");
                            rfnumber = rfnumber.Replace(".", "");
                            rfnumber = rfnumber.Replace(" ", "");
                            rfnumber = rfnumber.Replace("\t", "");
                            int.TryParse(rfnumber, out int rfnum);
                            if (rfnum > 0) { refnumbers.Add(rfnum); }
                        }
                    }
                    return GetEllidedNumbers(string.Join(",", refnumbers));
                }
                return null;
            }
            catch { return null; }
        }

        public static List<int> GetCitationsbyRange(string srngtext)
        {
            try
            {
                List<int> incitations = new List<int>();
                string stext = srngtext;
                stext = stext.Replace("\u2013", "-");
                stext = stext.Replace("\u2014", "-");
                stext = stext.Replace(", ", ",");
                stext = stext.Replace("; ", ",");
                stext = stext.Replace(" ", ",");
                stext = stext.Replace(",,", ",");
                stext = stext.Replace("(", "");
                stext = stext.Replace(")", "");
                stext = stext.Replace("[", "");
                stext = stext.Replace("]", "");
                stext = stext.Replace(".", "");
                foreach (string s in stext.Split(new char[] { ',' }))
                {
                    if (s.Contains("-"))
                    {
                        string[] ints = s.Split(new char[] { '-' });
                        if (ints.Length == 2)
                        {
                            int.TryParse(ints[0], out int num1);
                            int.TryParse(ints[1], out int num2);
                            if (num2 >= num1)
                            {
                                for (int i = num1; i <= num2; i++)
                                {
                                    incitations.Add(i);
                                }
                            }
                        }
                    }
                    else
                    {
                        int.TryParse(s, out int c);
                        if (c > 0) { incitations.Add(c); }
                    }
                }
                return incitations;
            }
            catch { return null; }
        }

    }
}
