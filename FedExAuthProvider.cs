using EmailGenerator.Interfaces;
using EmailGenerator.Models.Settings;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using Org.BouncyCastle.Asn1.Crmf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator
{
    public class FedExAuthProvider : IFedExAuthProvider
    {
        private readonly HttpClient _httpClient;
        private readonly FedExSettings _settings;

        public FedExAuthProvider(HttpClient httpClient, IOptions<FedExSettings> options)
        {
            _httpClient = httpClient;
            _settings = options.Value;
        }

        public async Task<string> GetAccessTokenAsync()
        {
            var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            { "grant_type", "client_credentials" },
            { "client_id", _settings.ClientId },
            { "client_secret", _settings.ClientSecret }
        });

            var response = await _httpClient.PostAsync("oauth/token", content);
            var responseContent = await response.Content.ReadAsStringAsync();

            response.EnsureSuccessStatusCode();
            dynamic result = JsonConvert.DeserializeObject(responseContent);
            return result.access_token;
        }
    }


}
