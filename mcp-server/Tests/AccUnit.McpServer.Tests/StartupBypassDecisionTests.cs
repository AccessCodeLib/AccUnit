using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class StartupBypassDecisionTests
    {
        [TestCase(true)]
        [TestCase(false)]
        [TestCase(null)]
        public void Create_BypassNotRequested_NoShiftKeyAndNoWarning(bool? allowBypassKey)
        {
            var decision = StartupBypassDecision.Create(false, allowBypassKey);

            Assert.That(decision.PressShiftKey, Is.False);
            Assert.That(decision.Warning, Is.Null);
        }

        [Test]
        public void Create_BypassAllowed_PressesShiftKeyWithoutWarning()
        {
            var decision = StartupBypassDecision.Create(true, true);

            Assert.That(decision.PressShiftKey, Is.True);
            Assert.That(decision.Warning, Is.Null);
        }

        [Test]
        public void Create_BypassKeyDisabled_NoShiftKeyAndWarning()
        {
            var decision = StartupBypassDecision.Create(true, false);

            Assert.That(decision.PressShiftKey, Is.False);
            Assert.That(decision.Warning, Does.Contain("AllowBypassKey = False"));
        }

        [Test]
        public void Create_BypassKeyUnknown_PressesShiftKeyAndWarns()
        {
            var decision = StartupBypassDecision.Create(true, null);

            Assert.That(decision.PressShiftKey, Is.True);
            Assert.That(decision.Warning, Does.Contain("could not be read"));
        }
    }
}
