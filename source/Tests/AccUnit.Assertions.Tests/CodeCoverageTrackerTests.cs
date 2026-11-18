using AccessCodeLib.AccUnit.CodeCoverage;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    public class CodeCoverageTrackerTests
    {
        // An empty tracker needs no VBProject: nothing has to be removed from a code module.

        [Test]
        public void Clear_WithoutArgument_ClearsAllModulesWithoutException()
        {
            var tracker = new CodeCoverageTracker(null);

            Assert.DoesNotThrow(() => tracker.Clear());
        }

        [Test]
        public void Clear_WithNull_ClearsAllModulesWithoutException()
        {
            var tracker = new CodeCoverageTracker(null);

            Assert.DoesNotThrow(() => tracker.Clear(null));
        }

        [Test]
        public void Clear_WithEmptyString_ClearsAllModulesWithoutException()
        {
            var tracker = new CodeCoverageTracker(null);

            Assert.DoesNotThrow(() => tracker.Clear(""));
        }
    }
}
