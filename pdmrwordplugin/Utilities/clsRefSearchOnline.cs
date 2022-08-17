using Newtonsoft.Json.Linq;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace pdmrwordplugin.Utilities
{
    public static class clsRefSearchOnline
    {
        private static HttpClient searchclient;
        public static void SetHttpClients()
        {
            searchclient = new HttpClient();
            searchclient.BaseAddress = new Uri("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/");
            System.Net.ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12 |
            SecurityProtocolType.Tls11 |
            SecurityProtocolType.Tls;
        }
        private static MediaTypeWithQualityHeaderValue newMediaTypeWithQualityHeaderValue(string v)
        {
            throw new NotImplementedException();
        }

        public static async Task<string> ISearchPubmed(ReferenceModel preference)
        {
            try
            {
                string sJsonOutput = await SearchPubmedOnline(preference);
                return sJsonOutput;
            }
            catch { return null; }
        }

        public static ReferenceModel ConsructPubQuery(ReferenceModel loadreference)
        {
            try
            {                
                JObject jobj = JObject.Parse(loadreference.RefJSON);
                ReferenceModel queryref = new ReferenceModel();
                queryref = loadreference;
                if (jobj.ContainsKey("author"))
                {
                    List<Author> jauthors = new List<Author>();
                    List<JToken> authorarray = jobj["author"].ToList();
                    foreach (JToken token in authorarray)
                    {
                        jauthors.Add(new Author()
                        {
                            given = token.Value<string>("given"),
                            family = token.Value<string>("family")
                        });
                    }
                    queryref.Authors = jauthors;
                }
                if (jobj.ContainsKey("volume"))
                {
                    queryref.volume = ((JValue)((JContainer)jobj["volume"]).First).Value.ToString();
                }
                if (jobj.ContainsKey("date"))
                {
                    queryref.date = ((JValue)((JContainer)jobj["date"]).First).Value.ToString();
                }
                if (jobj.ContainsKey("issue"))
                {
                    queryref.issue = ((JValue)((JContainer)jobj["issue"]).First).Value.ToString();
                }
                if (jobj.ContainsKey("pages"))
                {
                    queryref.pages = ((JValue)((JContainer)jobj["pages"]).First).Value.ToString();
                }
                if (jobj.ContainsKey("title"))
                {
                    queryref.title = ((JValue)((JContainer)jobj["title"]).First).Value.ToString();
                }
                if (jobj.ContainsKey("container-title"))
                    queryref.containertitle = ((JValue)((JContainer)jobj["container-title"]).First).Value.ToString();                
                return queryref;
            }
            catch { return loadreference; }
        }

        public static async Task<string> FetchPubmedContent(string pubids)
        {
            try
            {
                string resultJson = "";
                string sQuery = "efetch.fcgi?db=pubmed&id=<idlist>&retmode=xml";
                sQuery = sQuery.Replace("<idlist>", pubids);
                HttpResponseMessage response = await searchclient.GetAsync(sQuery);
                if (response.IsSuccessStatusCode)
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }
                return resultJson;
            }
            catch { return ""; }
        }

        public static string GetBetterMatchRec(string pubxml, ReferenceModel cReference)
        {
            try
            {
                string matchedxml = "";
                var xDoc = XDocument.Parse(pubxml);
                var articles = (from item in xDoc.Descendants("PubmedArticle") select item).ToList();
                foreach (var article in articles)
                {
                    string orgreftext = cReference.Reftext;
                    string matchedelems = "";
                    double resmatch = 0;
                    var authors = from item in article.Descendants("Author") select item;
                    foreach (var author in authors)
                    {
                        var lastname = author.Element("LastName");
                        if (lastname != null)
                        {
                            if (orgreftext.Contains(lastname.Value)) { matchedelems += "$AuthorFound$"; }
                            else if (cReference.Authors != null && cReference.Authors.Count > 0)
                            {
                                resmatch = Functions.ClsCommonUtils.CalculateSimilarity(lastname.Value, cReference.Authors.FirstOrDefault().family);
                                if (resmatch > 0.8) { matchedelems += "$AuthorFound$"; }
                            }
                        }
                    }
                    var pubdate = from item in article.Descendants("PubDate") select item;
                    if (pubdate != null && pubdate.Count() > 0)
                    {
                        var date = pubdate.FirstOrDefault().Element("Year");
                        //check for epub date too
                        if (date != null)
                        {                            
                            if (orgreftext.Contains(date.Value))
                            {
                                matchedelems += "$YearFound$";
                                //orgreftext = orgreftext.Replace(date.Value, "");
                            }
                        }
                    }
                    var pagination = from item in article.Descendants("MedlinePgn") select item;
                    if (pagination != null && pagination.Count() > 0)
                    {
                        string firstpagenum = GetFirstPagenumber(pagination.FirstOrDefault().Value);
                        if (orgreftext.Contains(firstpagenum + "-") || orgreftext.Contains(firstpagenum + "\u2014") || orgreftext.ToLower().Contains("p" + firstpagenum) 
                            || orgreftext.ToLower().Contains("p. " + firstpagenum))
                        {
                            matchedelems += "$PageFound$";
                            //orgreftext = Regex.Replace(orgreftext, firstpagenum + @"[\u2014\-]\d+", "");
                        }
                    }
                    var volume = from item in article.Descendants("Volume") select item;
                    if (volume != null && volume.Count() > 0)
                    {
                        if (orgreftext.Contains(volume.FirstOrDefault().Value))
                        {
                            matchedelems += "$VolumeFound$";
                            //orgreftext = orgreftext.Replace(volume.FirstOrDefault().Value, "");
                        }
                    }
                    var title = from item in article.Descendants("ArticleTitle") select item;
                    if (title != null && title.Count() > 0)
                    {
                        if (orgreftext.ToLower().Contains(title.FirstOrDefault().Value.ToLower()))
                        {
                            matchedelems += "$TitleFound$";
                        }
                        else
                        {
                            resmatch = Functions.ClsCommonUtils.CalculateSimilarity(title.FirstOrDefault().Value, cReference.title);
                            if (resmatch > 0.9) { matchedelems += "$TitleFound$"; }
                        }
                    }
                    var jtitle = from item in article.Descendants("MedlineTA") select item;
                    if (jtitle != null && jtitle.Count() > 0)
                    {
                        if (orgreftext.ToLower().Contains(jtitle.FirstOrDefault().Value.ToLower()) ||
                            orgreftext.Replace(".", "").ToLower().Contains(jtitle.FirstOrDefault().Value.ToLower()))
                        {
                            matchedelems += "$JTitleFound$";
                        }
                        else
                        {
                            resmatch = Functions.ClsCommonUtils.CalculateSimilarity(jtitle.FirstOrDefault().Value, cReference.containertitle);
                            if (resmatch > 0.9) { matchedelems += "$JTitleFound$"; }
                        }
                    }
                    jtitle = from item in article.Descendants("Title") select item;
                    if (jtitle != null && jtitle.Count() > 0)
                    {
                        if (orgreftext.ToLower().Contains(jtitle.FirstOrDefault().Value.ToLower()))
                        {
                            matchedelems += "$JTitleFound$";
                        }
                        else
                        {
                            resmatch = Functions.ClsCommonUtils.CalculateSimilarity(jtitle.FirstOrDefault().Value, cReference.containertitle);
                            if (resmatch > 0.9) { matchedelems += "$JTitleFound$"; }
                        }
                    }
                    jtitle = from item in article.Descendants("ISOAbbreviation") select item;
                    if (jtitle != null && jtitle.Count() > 0)
                    {
                        if (orgreftext.ToLower().Contains(jtitle.FirstOrDefault().Value.ToLower()) ||
                            orgreftext.Replace(".", "").ToLower().Contains(jtitle.FirstOrDefault().Value.ToLower()))
                        {
                            matchedelems += "$JTitleFound$";
                        }
                        else
                        {
                            resmatch = Functions.ClsCommonUtils.CalculateSimilarity(jtitle.FirstOrDefault().Value, cReference.containertitle);
                            if (resmatch > 0.9) { matchedelems += "$JTitleFound$"; }
                        }
                    }
                    // Check the conditions
                    if (matchedelems.Contains("$AuthorFound$") && matchedelems.Contains("$YearFound$") && matchedelems.Contains("$PageFound$") &&
                        matchedelems.Contains("$VolumeFound$"))
                    {
                        matchedxml= article.ToString();
                    }
                    if (matchedelems.Contains("$AuthorFound$") && matchedelems.Contains("$TitleFound$") && matchedelems.Contains("$PageFound$") &&
                        matchedelems.Contains("$JTitleFound$"))
                    {
                        matchedxml = article.ToString();
                    }
                    if (matchedelems.Contains("$AuthorFound$") && matchedelems.Contains("$YearFound$") && matchedelems.Contains("$PageFound$") &&
                        matchedelems.Contains("$JTitleFound$"))
                    {
                        matchedxml = article.ToString();
                    }
                    if (matchedelems.Contains("$AuthorFound$") && matchedelems.Contains("$YearFound$") && matchedelems.Contains("$TitleFound$"))
                    {
                        matchedxml = article.ToString();
                    }
                    if (!string.IsNullOrEmpty(matchedxml)) { break; }
                    // Ends here
                }
                return matchedxml;
            }
            catch { return ""; }
        }

        private static string GetFirstPagenumber(string strpage)
        {
            try
            {
                if (Regex.IsMatch(strpage, @"\d+"))
                {
                    return Regex.Matches(strpage, @"\d+")[0].Value;
                }
                return strpage;
            }
            catch { return strpage; }
        }

        public static async Task<string> SearchPubmedOnline(ReferenceModel sreference)
        {
            try
            {
                ReferenceModel prcReference = ConsructPubQuery(sreference);
                string[] queries = "<author>[au]+AND+<date>[dp]+AND+<volume>[vi]#<author>[au]+AND+<date>[dp]+AND+<pagination>[pg]#<author>[au]+AND+<date>[dp]+AND+<containertitle>[ta]#<title>#<title>+AND+<date>[dp]+AND+<volume>[vi]#<author>[au]+AND+<date>[dp]+AND+<containertitle>[jt]#<author>[au]+AND+<date>[dp]<volume>[vi]+AND+<date>[dp]+AND+<containertitle>[jt]#<volume>[vi]+AND+<date>[dp]+AND+<containertitle>[ta]".Split('#');
                string resultJson = "";
                foreach (string query in queries)
                {
                    string sQuery = "esearch.fcgi?db=pubmed&usehistory=y&retmax=25&retmode=json";
                    string strtemp = query;
                    if (prcReference.Authors != null && prcReference.Authors.Count > 0)
                    {   
                        List<Author> jauthors = new List<Author>(prcReference.Authors);
                        if (query.Contains("<author>"))
                            strtemp = strtemp.Replace("<author>", prcReference.Authors.FirstOrDefault().family);
                        if (query.Contains("<date>"))
                            strtemp = strtemp.Replace("<date>", prcReference.date);
                        if (query.Contains("<volume>") && !string.IsNullOrEmpty(prcReference.volume))
                            strtemp = strtemp.Replace("<volume>", prcReference.volume);
                        if (query.Contains("<containertitle>") && !string.IsNullOrEmpty(prcReference.containertitle))
                            strtemp = strtemp.Replace("<containertitle>", prcReference.containertitle);
                        if (query.Contains("<title>") && !string.IsNullOrEmpty(prcReference.title))
                            strtemp = strtemp.Replace("<title>", prcReference.title);
                        if (query.Contains("<pagination>") && !string.IsNullOrEmpty(prcReference.pages))
                            strtemp = strtemp.Replace("<pagination>", prcReference.pages);
                        sQuery += "&term=" + strtemp;
                        HttpResponseMessage response = await searchclient.GetAsync(sQuery);
                        if (response.IsSuccessStatusCode)
                        {
                            resultJson = await response.Content.ReadAsStringAsync();
                            if (!string.IsNullOrEmpty(resultJson))
                            {
                                JObject jobj = JObject.Parse(resultJson);
                                if (jobj.ContainsKey("esearchresult"))
                                {   
                                    var refidlists = (from p in jobj["esearchresult"]["idlist"] select (string)p).ToList();
                                    if (refidlists.Count > 0)
                                    {
                                        string pubmedids = string.Join(",", refidlists);
                                        resultJson = await FetchPubmedContent(pubmedids);
                                        string sBettermatch = GetBetterMatchRec(resultJson, prcReference);
                                        if (!string.IsNullOrEmpty(sBettermatch)) { resultJson = sBettermatch; break; }
                                    }
                                }
                            }
                        }
                    }
                }
                return resultJson;
            }
            catch(Exception ex) { return ex.Message; }
        }
    }
}
