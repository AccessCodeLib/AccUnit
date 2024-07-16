using NUnit.Framework.Constraints;
using AccessCodeLib.AccUnit.Extension.OpenAI;

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
            const string key = "xyz";
            Environment.SetEnvironmentVariable("OPENAI_API_KEY", key);

            var service = new OpenAiService();
            var apiKey = service.ApiKey;

            Assert.That(apiKey, Is.EqualTo(key));
        }

        [Test]
        public void ReadApiKey()
        {
            var service = new OpenAiService();
            var apiKey = service.ApiKey;

            Assert.That(apiKey, Is.GreaterThan(""), "please check UserSecrets config");
        }

    }
}