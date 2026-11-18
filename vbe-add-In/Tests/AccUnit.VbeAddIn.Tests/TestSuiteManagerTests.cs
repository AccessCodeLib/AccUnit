using NUnit.Framework;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn.Tests
{
    public class TestSuiteManagerTests
    {
        [Test]
        public void TestSuite_WhenInitializationFails_ThrowsWithOriginalMessageInsteadOfReturningNull()
        {
            using (var manager = new TestSuiteManager())
            {
                manager.TestResultReporterRequest += (ref System.Collections.Generic.IList<Interfaces.ITestResultReporter> reporters) =>
                    throw new InvalidOperationException("reporter failure");
                // no OfficeApplicationHelper => the suite cannot be created

                var exception = Assert.Throws<InvalidOperationException>(() => { var _ = manager.TestSuite; });

                Assert.That(exception.Message, Does.StartWith("AccUnit test suite could not be initialized"));
                Assert.That(exception.InnerException, Is.Not.Null);
                Assert.That(exception.Message, Does.Contain(exception.InnerException.Message));
            }
        }
    }
}
