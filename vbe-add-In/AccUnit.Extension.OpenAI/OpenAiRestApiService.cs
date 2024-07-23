using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenAI_API.Models;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiRestApiService
    {
        const string _apiUrl = "https://api.openai.com/v1/chat/completions";
        private static readonly HttpClient _client = new HttpClient();
        private readonly string _apiKey;

        public OpenAiRestApiService(string apiKey)
        {
            _apiKey = apiKey;
        }

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
