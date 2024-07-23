using System;
using Microsoft.Extensions.Configuration;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiService
    {
        string SendRequest(object[] messages, int maxToken = 500, string model = null); 
        string Model { get; set; }
    }

    public class OpenAiService : IOpenAiService
    {
        public const string CredentialKey = "AccessCodeLib.OpenAI.ApiKey";
        public const string EnvironmentKey = "OPENAI_API_KEY";

        private string _apiKey;
        private readonly ICredentialManager _credentialManager;
        private readonly OpenAiRestApiService _restService;

        public OpenAiService(ICredentialManager credentialManager, string gptModel = "gpt-4o-mini")
        {
            _credentialManager = credentialManager;
            Model = gptModel;
            _restService = new OpenAiRestApiService(ApiKey);
        }

        public string Model { get; set; }

        public string ApiKey
        {
            get
            {
                if (string.IsNullOrEmpty(_apiKey))
                {
                    _apiKey = GetApiKey();
                    Console.WriteLine("api key: " + _apiKey);
                }
                return _apiKey;
            }
        }

        #region OpenAI API Key
        private string GetApiKey()
        {
            string apiKey;
       
            // only for debug or test?
            apiKey = GetApiKeyFromEnvironment();
            if (!string.IsNullOrEmpty(apiKey))
            {
                return apiKey;
            }
            
            apiKey = GetApiKeyFromUserSecrets();
            if (!string.IsNullOrEmpty(apiKey))
            {
                return apiKey;
            }
           
            apiKey = GetApiKeyFromCredentialManager();
            if (!string.IsNullOrEmpty(apiKey))
            {
               return apiKey;
            }

            return null;
        }

        private string GetApiKeyFromCredentialManager()
        {
            return _credentialManager.Retrieve(CredentialKey);
        }

        private string GetApiKeyFromUserSecrets()
        {
            try
            {
                var configuration = new ConfigurationBuilder()
                .AddUserSecrets("AccessCodeLib.OpenAI.UserSecrets")
                .Build();
                return configuration[CredentialKey];
            }
            catch
            {
                return null;
            }
            
        }

        private string GetApiKeyFromEnvironment()
        {
            return Environment.GetEnvironmentVariable(EnvironmentKey);
        }
        

        public void StoreApiKey(string apiKey)
        {
            _apiKey = apiKey;
            string username = Environment.UserName;
            _credentialManager.Save(CredentialKey, username, apiKey);
        }
        #endregion

        public string SendRequest(object[] messages, int maxToken = 500, string model = null)
        {
            if (!string.IsNullOrEmpty(model))
            {
                Model = model;
            }
            Console.WriteLine(ApiKey);

            var requestBody = new
            {
                model = Model,
                messages,
                max_tokens = maxToken,
                temperature = 0.2  
            };

            var jsonRequestBody = JsonConvert.SerializeObject(requestBody);

            //Console.WriteLine(jsonRequestBody.Replace(@"\r\n", "\r\n"));

            return _restService.SendRequest(jsonRequestBody);
        }
    }
}
