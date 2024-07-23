using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiRestApiService
    {
        const string _apiUrl = "https://api.openai.com/v1/chat/completions";
        private readonly HttpClient _client;
        private readonly string _apiKey;

        public OpenAiRestApiService(string apiKey)
        {
            _apiKey = apiKey;

            // Fügen Sie diesen Code hinzu, bevor Sie eine HttpClient-Instanz erstellen
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 |
                                                             System.Net.SecurityProtocolType.Tls13;


            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => { return true; }
            };
            _client = new HttpClient(handler);
        }

        public string SendRequest(string jsonRequestBody)
        {
            
            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_apiUrl),
                Headers =
            {
                { "Authorization", $"Bearer {_apiKey}" }
            },
                Content = new StringContent(jsonRequestBody, Encoding.UTF8, "application/json")
            };

            HttpResponseMessage response = _client.SendAsync(request).Result;
            response.EnsureSuccessStatusCode();

            string responseBody = response.Content.ReadAsStringAsync().Result;

            var jsonResponse = JObject.Parse(responseBody);
            var choicesContent = jsonResponse["choices"][0]["message"]["content"].ToString();

            return choicesContent;

        }
    }
}
