﻿using Newtonsoft.Json.Linq;
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

        public static string GetBetterMatchRec(string pubxml, string orgreftext)
        {
            try
            {
                string matchedelems = "";
                var xDoc = XDocument.Parse(pubxml);
                var articles = from item in xDoc.Descendants("PubmedArticle") select item;
                foreach (var article in articles)
                {
                    var authors = from item in article.Descendants("Author") select item;
                    foreach (var author in authors)
                    {
                        var lastname = author.Element("LastName");
                        if (lastname != null)
                        {
                            if (orgreftext.Contains(lastname.Value)) { matchedelems += "$Authorfound$"; }
                        }
                    }
                    var pubdate = from item in article.Descendants("PubDate") select item;
                    if (pubdate != null)
                    {
                        var date = pubdate.FirstOrDefault().Element("Year");
                        if (date != null)
                        {
                            if (orgreftext.Contains(date.Value))
                            {
                                matchedelems += "$YearFound$";
                                orgreftext = orgreftext.Replace(date.Value, "");
                            }
                        }
                    }
                    var pagination = from item in article.Descendants("MedlinePgn") select item;
                    if (pagination != null)
                    {
                        string firstpagenum = GetFirstPagenumber(pagination.FirstOrDefault().Value);
                        if (orgreftext.Contains(firstpagenum + "-") || orgreftext.Contains(firstpagenum + "\u2014"))
                        {
                            matchedelems += "$PageFound$";
                            orgreftext = Regex.Replace(orgreftext, firstpagenum + @"[\u2014\-]\d+", "");
                        }
                    }
                    var volume = from item in article.Descendants("Volume") select item;
                    if (volume != null)
                    {
                        if (orgreftext.Contains(volume.FirstOrDefault().Value))
                        {
                            matchedelems += "$VolumeFound$";
                            orgreftext = orgreftext.Replace(volume.FirstOrDefault().Value, "");
                        }
                    }
                }
                return "";
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
                string[] queries = "<author>[au]+AND+<date>[dp]+AND+<volume>[volume]#<author>[au]+AND+<date>[dp]#<author>[au]+AND+<date>[dp]+AND+<containertitle>[ta]".Split('#');
                string resultJson = "";
                foreach (string query in queries)
                {
                    string sQuery = "esearch.fcgi?db=pubmed&usehistory=y&retmax=10&retmode=json";
                    string strtemp = query;
                    if (prcReference.Authors != null && prcReference.Authors.Count > 0)
                    {   
                        List<Author> jauthors = new List<Author>(prcReference.Authors);
                        if (query.Contains("<author>"))
                            strtemp = strtemp.Replace("<author>", prcReference.Authors.FirstOrDefault().family);
                        if (query.Contains("<date>"))
                            strtemp = strtemp.Replace("<date>", prcReference.date);
                        if (query.Contains("<volume>"))
                            strtemp = strtemp.Replace("<volume>", prcReference.volume);
                        if (query.Contains("<containertitle>"))
                            strtemp = strtemp.Replace("<containertitle>", prcReference.containertitle);
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
                                        string sBettermatch = GetBetterMatchRec(resultJson, prcReference.Reftext);
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
