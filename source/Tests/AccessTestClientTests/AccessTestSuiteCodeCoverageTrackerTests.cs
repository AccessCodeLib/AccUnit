using AccessCodeLib.AccUnit.CodeCoverage;
using NUnit.Framework;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class AccessTestSuiteCodeCoverageTrackerTests
    {
        private class TestableAccessTestSuite : Interop.AccessTestSuite
        {
            public TestableAccessTestSuite() : base(null, null, null, null)
            {
            }

            public void Trace(string text)
            {
                RaiseTraceMessage(text);
            }
        }

        [Test]
        public void CodeCoverageTracker_SetViaInteropInterface_IsPassedToTraceMessageEvent()
        {
            ICodeCoverageTracker received = null;
            var tracker = new Interop.CodeCoverageTracker();
            using (var suite = new TestableAccessTestSuite())
            {
                ((Interop.IAccessTestSuite)suite).CodeCoverageTracker = tracker;
                suite.TestTraceMessage += (message, codeCoverageTracker) => received = codeCoverageTracker;

                suite.Trace("Total");
            }

            Assert.That(received, Is.SameAs(tracker));
        }

        [Test]
        public void CodeCoverageTracker_NotSet_ReturnsNull()
        {
            using (var suite = new TestableAccessTestSuite())
            {
                Assert.That(((Interop.IAccessTestSuite)suite).CodeCoverageTracker, Is.Null);
            }
        }

        [Test]
        public void IAccessTestSuite_CodeCoverageTracker_IsComVisible()
        {
            var property = typeof(Interop.IAccessTestSuite).GetProperty(nameof(Interop.IAccessTestSuite.CodeCoverageTracker),
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

            Assert.That(property, Is.Not.Null, "Property must be declared on IAccessTestSuite");
            Assert.That(property.CanRead && property.CanWrite, Is.True);
            Assert.That(property.GetCustomAttributes(typeof(ComVisibleAttribute), true)
                .Cast<ComVisibleAttribute>().Any(a => !a.Value), Is.False);
        }
    }
}
