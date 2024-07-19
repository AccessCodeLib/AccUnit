using System;
using Microsoft.Extensions.Configuration;
using OpenAI.Chat;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface IOpenAiService
    {
        ChatClient NewChatClient(string model = null);
    }

    public class OpenAiService : IOpenAiService
    {
        public const string CredentialKey = "AccessCodeLib.OpenAI.ApiKey";
        public const string EnvironmentKey = "OPENAI_API_KEY";

        private string _apiKey;
        private readonly ICredentialManager _credentialManager;

        private string _gptModel;

        public OpenAiService(ICredentialManager credentialManager, string gptModel = "gpt-4o")
        {
            _credentialManager = credentialManager;
            _gptModel = gptModel;
        }

        public string Model
        {
            get => _gptModel;
            set => _gptModel = value;
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
            var configuration = new ConfigurationBuilder()
                .AddUserSecrets("AccessCodeLib.OpenAI.UserSecrets")
                .Build();
            return configuration[CredentialKey];
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

        public ChatClient NewChatClient(string model = null)
        {
            if (string.IsNullOrEmpty(model))
            {
                model = _gptModel;
            }
            Console.WriteLine(ApiKey);
            var apiKeyCredential = new System.ClientModel.ApiKeyCredential(ApiKey);
            return new ChatClient(model: model, credential: apiKeyCredential);
        }
    }
}
