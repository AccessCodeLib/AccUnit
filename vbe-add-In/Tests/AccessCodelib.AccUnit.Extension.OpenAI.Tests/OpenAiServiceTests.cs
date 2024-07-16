using NUnit.Framework.Constraints;
using AccessCodeLib.AccUnit.Extension.OpenAI;
using OpenAI.Chat;

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

            Assert.That(apiKey, Is.EqualTo(secretKey));
        }

        [Test]
        public void ReadApiKeyFromUserSecrets()
        {
            var service = new OpenAiService(new TestSupport.CredentialManagerMock());
            var apiKey = service.ApiKey;
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

            ChatCompletion chatCompletion = client.CompleteChat(
            [
                new UserChatMessage("Say 'This is a test.'"),
            ]);

            var actual = chatCompletion.Content[0].Text; 

            Assert.That(actual, Is.EqualTo("This is a test."));
        }

    }
}