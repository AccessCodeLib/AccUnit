using NUnit.Framework.Constraints;
using AccessCodeLib.AccUnit.Extension.OpenAI;
using OpenAI_API.Chat;
using OpenAI_API.Models;

namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests
{
    public class OpenAiServiceTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void ReadApiKeyFromEnvironment()
        {
            var secretKey = Random.Shared.GetHashCode().ToString();
            Environment.SetEnvironmentVariable("OPENAI_API_KEY", secretKey);

            var service = new OpenAiService(new CredentialManager());
            var apiKey = service.ApiKey;

            Environment.SetEnvironmentVariable("OPENAI_API_KEY", string.Empty);

            Assert.That(apiKey, Is.EqualTo(secretKey));
        }

        [Test]
        public void ReadApiKeyFromUserSecrets()
        {
            var service = new OpenAiService(new TestSupport.CredentialManagerMock());
            var apiKey = service.ApiKey;
            Console.WriteLine(apiKey);  
            Assert.That(apiKey, Is.Not.Null, "please check UserSecrets config");
            Assert.That(apiKey, Is.GreaterThan(""), "please check UserSecrets config");
        }

        [Test]
        public void StoreApiKey()
        {
            var secretKey = Random.Shared.GetHashCode().ToString();
            var credentialManager = new TestSupport.CredentialManagerMock();

            var service = new OpenAiService(credentialManager);
            service.StoreApiKey(secretKey);

            var actual = credentialManager.Retrieve(OpenAiService.CredentialKey);
            Assert.That(actual, Is.EqualTo(secretKey));
        }

        [Test]
        public void GetChatClient()
        {
            var service = new OpenAiService(new CredentialManager());
            var client = service.NewChatClient();

            var messages = new List<ChatMessage>();
            messages.Add(new ChatMessage(ChatMessageRole.User, "Say 'This is a test.'"));

            var request = new ChatRequest()
            {
                Model = new Model("gpt-4o"), // Model.ChatGPTTurbo,
                Temperature = 0.1,
                MaxTokens = 5,
                Messages = messages
            };
           
            var result = client.CreateChatCompletionAsync(request).Result;
            var actual = result.ToString();

            Assert.That(actual, Is.EqualTo("This is a test."));
        }

    }
}