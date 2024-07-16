using System;
using Microsoft.Extensions.Configuration;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiService
    {
        public const string CredentialKey = "AccessCodeLib.OpenAI.ApiKey";
        public const string EnvironmentKey = "OPENAI_API_KEY";

        private string _apiKey;
        private readonly ICredentialManager _credentialManager;

        public OpenAiService(ICredentialManager credentialManager)
        {
            _credentialManager = credentialManager;
        }

        public string ApiKey
        {
            get
            {
                if (string.IsNullOrEmpty(_apiKey))
                {
                    _apiKey = GetApiKey();
                }
                return _apiKey;
            }
        }

        #region ReadApiKey
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
            var configuration = new ConfigurationBuilder()
                .AddUserSecrets("AccessCodeLib.OpenAI.UserSecrets")
                .Build();
            return configuration[CredentialKey];
        }

        private string GetApiKeyFromEnvironment()
        {
            return Environment.GetEnvironmentVariable(EnvironmentKey);
        }
        #endregion

        public void StoreApiKey(string apiKey)
        {
            _apiKey = apiKey;
            string username = Environment.UserName;
            _credentialManager.Save(CredentialKey, username, apiKey);
        }
    }
}
