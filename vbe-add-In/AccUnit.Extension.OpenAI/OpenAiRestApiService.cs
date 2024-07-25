using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiRestApiService : IOpenAiRestApiService
    {
        const string _apiUrl = "https://api.openai.com/v1/chat/completions";
        private string _apiKey;
        private readonly HttpClient _client;

        public OpenAiRestApiService()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 |
                                                              System.Net.SecurityProtocolType.Tls13;

            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => { return true; }
            };
            _client = new HttpClient(handler);
        }

        public string ApiKey { get => _apiKey; set => _apiKey = value; }

        public async Task<string> SendRequest(string jsonRequestBody)
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

            HttpResponseMessage response = await _client.SendAsync(request);
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();
            var jsonResponse = JObject.Parse(responseBody);
            var choicesContent = jsonResponse["choices"][0]["message"]["content"].ToString();

            return choicesContent;
        }
    }
}
