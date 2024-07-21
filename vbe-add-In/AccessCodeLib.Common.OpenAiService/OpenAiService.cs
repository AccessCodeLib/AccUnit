using System;
//using Microsoft.Extensions.Configuration;
using OpenAI_API;
using OpenAI_API.Chat;
using OpenAI_API.Models;


namespace AccessCodeLib.Common.OpenAI
{
    public interface IOpenAiService
    {
        OpenAI_API.Chat.IChatEndpoint NewChatClient(string model = null);
        Model Model { get; set; }  
    }

    public class OpenAiService : IOpenAiService
    {
        public const string CredentialKey = "AccessCodeLib.OpenAI.ApiKey";
        public const string EnvironmentKey = "OPENAI_API_KEY";

        private string _apiKey;
        private readonly ICredentialManager _credentialManager;

        public Model Model { get; set; }

        private readonly OpenAIAPI _api;

        public OpenAiService(ICredentialManager credentialManager, string gptModel = "gpt-4o")
        {
            _credentialManager = credentialManager;
            Model = new Model(gptModel);
            _api = new OpenAIAPI(ApiKey);
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
            
            /*
            apiKey = GetApiKeyFromUserSecrets();
            if (!string.IsNullOrEmpty(apiKey))
            {
                return apiKey;
            }
            */

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

        /*
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
        */

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

        public IChatEndpoint NewChatClient(string model = null)
        {
            if (!string.IsNullOrEmpty(model))
            {
                Model = new Model(model);
            }
            Console.WriteLine(ApiKey);

            return _api.Chat;

            //var apiKeyCredential = new System.ClientModel.ApiKeyCredential(ApiKey);
            //return new ChatClient(model: model, credential: apiKeyCredential);
        }
    }
}
