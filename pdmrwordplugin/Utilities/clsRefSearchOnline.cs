using Newtonsoft.Json.Linq;
using pdmrwordplugin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

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
                                    int resultcount = JObject.Parse(resultJson)["esearchresult"].Count();
                                    if (resultcount > 0)
                                    {
                                        resultJson = "";
                                        //break;
                                    }
                                    //Check the conditions and break;
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
