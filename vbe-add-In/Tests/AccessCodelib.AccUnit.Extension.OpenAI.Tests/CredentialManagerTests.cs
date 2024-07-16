namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests
{
    public class CredentialManagerTests
    {
        [Test]
        public void StoreApiKey()
        {
            var cm = new CredentialManager();
            const string secretToStore = "abc";
            const string targetKey = "AccessCodeLib.Test.SecretKey";

            cm.Save(targetKey, "username123", secretToStore);
            var actual = cm.Retrieve(targetKey);

            Assert.That(actual, Is.EqualTo(secretToStore));
        }
    }
}
