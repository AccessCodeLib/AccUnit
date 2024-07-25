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

            var restService = new OpenAiRestApiService();   
            var service = new OpenAiService(new CredentialManager(), restService);
            var apiKey = restService.ApiKey;

            Environment.SetEnvironmentVariable("OPENAI_API_KEY", string.Empty);

            Assert.That(apiKey, Is.EqualTo(secretKey));
        }

        [Test]
        public void ReadApiKeyFromUserSecrets()
        {
            var restService = new OpenAiRestApiService();
            var service = new OpenAiService(new TestSupport.CredentialManagerMock(), restService);
            var apiKey = restService.ApiKey;
            Console.WriteLine(apiKey);  
            Assert.That(apiKey, Is.Not.Null, "please check UserSecrets config");
            Assert.That(apiKey, Is.GreaterThan(""), "please check UserSecrets config");
        }

        [Test]
        public void StoreApiKey()
        {
            var secretKey = Random.Shared.GetHashCode().ToString();
            var credentialManager = new TestSupport.CredentialManagerMock();

            var service = new OpenAiService(credentialManager, new OpenAiRestApiService());
            service.StoreApiKey(secretKey);

            var actual = credentialManager.Retrieve(OpenAiService.CredentialKey);
            Assert.That(actual, Is.EqualTo(secretKey));
        }

        [Test]
        public void GetChatClient()
        {
            var service = new OpenAiService(new CredentialManager(), new OpenAiRestApiService() );

            var messages = new []
            {
                new { role = "system", content = "Say 'This is a test.'" }
            };

            var result = service.SendRequest(messages);
            var actual = result.ToString();

            Assert.That(actual, Is.EqualTo("This is a test."));
        }

    }
}