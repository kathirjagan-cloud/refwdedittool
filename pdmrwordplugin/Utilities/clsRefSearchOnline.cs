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

        public static async Task<string> SearchPubmedOnline(ReferenceModel sreference)
        {
            try
            {
                string sQuery = "esearch.fcgi?db=pubmed&term=cancer&usehistory=y&retmax=10&retmode=json";
                HttpResponseMessage response = await searchclient.GetAsync(sQuery);
                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                    return null;
            }
            catch(Exception ex) { return ex.Message; }
        }
    }
}
