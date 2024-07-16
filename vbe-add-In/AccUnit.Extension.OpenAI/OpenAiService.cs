using System;
using Microsoft.Extensions.Configuration;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class OpenAiService
    {
        private const string CredentialKey = "AccessCodeLib:OpenAI.ApiKey";
        private const string EnvironmentKey = "OPENAI_API_KEY";

        private string _apiKey;

        public OpenAiService()
        {
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

            apiKey = GetApiKeyFromWindowsCredential();
            if (!string.IsNullOrEmpty(apiKey))
            {
               return apiKey;
            }

            return null;
        }

        private string GetApiKeyFromWindowsCredential()
        {
            var cm = new CredentialManager();
            return cm.Retrieve(CredentialKey);
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

        public void SetApiKey(string apiKey)
        {
            _apiKey = apiKey;
            var cm = new CredentialManager();
            string username = Environment.UserName;

            cm.Save("CredentialKey", username, apiKey);
        }
    }
}
