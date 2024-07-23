using System;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiService : IOpenAiService
    {
        public const string CredentialKey = "AccessCodeLib.OpenAI.ApiKey";
        public const string EnvironmentKey = "OPENAI_API_KEY";

        private readonly ICredentialManager _credentialManager;
        private readonly IOpenAiRestApiService _restService;

        public OpenAiService(ICredentialManager credentialManager, IOpenAiRestApiService openAiRestApiService,  int maxToken = 250, string gptModel = "gpt-4o-mini")
        {
            Model = gptModel;
            MaxToken = maxToken;

            _credentialManager = credentialManager;
            _restService = openAiRestApiService;
            _restService.ApiKey = GetApiKey();
        }

        public string Model { get; set; }
        public int MaxToken { get; set; }   

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
            _restService.ApiKey = apiKey;   
            string username = Environment.UserName;
            _credentialManager.Save(CredentialKey, username, apiKey);
        }
        #endregion

        public string SendRequest(object[] messages, int maxToken = 0, string model = null)
        {
            if (string.IsNullOrEmpty(model))
            {
                model = Model;
            }
            
            if (maxToken == 0)
            {
                maxToken = MaxToken;
            }

            var requestBody = new
            {
                model,
                messages,
                max_tokens = maxToken,
                temperature = 0.2  
            };

            var jsonRequestBody = JsonConvert.SerializeObject(requestBody);
            return _restService.SendRequest(jsonRequestBody);
        }
    }
}
