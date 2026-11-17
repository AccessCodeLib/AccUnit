using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class AccessInstanceOwnershipTests
    {
        [Test]
        public void IsStartedByServer_NewProcess_IsTrue()
        {
            Assert.That(AccessInstanceOwnership.IsStartedByServer(300, new[] { 100, 200 }), Is.True);
        }

        [Test]
        public void IsStartedByServer_NoAccessRunningBefore_IsTrue()
        {
            Assert.That(AccessInstanceOwnership.IsStartedByServer(300, new int[0]), Is.True);
        }

        [Test]
        public void IsStartedByServer_ProcessRunningBefore_IsFalse()
        {
            Assert.That(AccessInstanceOwnership.IsStartedByServer(200, new[] { 100, 200 }), Is.False);
        }
    }
}
