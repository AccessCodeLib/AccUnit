using System.Linq;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class JUnitResultReaderTests
    {
        // Format of AccessCodeLib.AccUnit.Integration.JUnitXmlWriter
        private const string JUnitXml =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<testsuites name=""AccUnit"" tests=""5"" failures=""1"" errors=""1"" skipped=""1"" time=""0.123"" timestamp=""2026-11-15T22:00:00"">
  <testsuite name=""clsSampleTests"" tests=""5"" failures=""1"" errors=""1"" skipped=""1"" time=""0.123"" timestamp=""2026-11-15T22:00:00"">
    <testcase name=""PassingTest"" classname=""clsSampleTests"" time=""0.012"" />
    <testcase name=""FailingTest"" classname=""clsSampleTests"" time=""0.003"">
      <failure message=""actual is less than expected (Expected: 2 but was: 1)"">actual is less than expected (Expected: 2 but was: 1)</failure>
    </testcase>
    <testcase name=""ErrorTest"" classname=""clsSampleTests"" time=""0.001"">
      <error message=""runtime error in test"">runtime error in test</error>
    </testcase>
    <testcase name=""IgnoredTest"" classname=""clsSampleTests"" time=""0"">
      <skipped message="""" />
    </testcase>
    <testcase name=""RowTest.NotEqual"" classname=""clsSampleTests"" time=""0.002"">
      <failure message=""actual is less than expected (Expected: 3 but was: 2)"">actual is less than expected (Expected: 3 but was: 2)</failure>
    </testcase>
  </testsuite>
</testsuites>";

        [Test]
        public void Read_ReturnsOneResultPerTestCase()
        {
            var results = JUnitResultReader.Read(JUnitXml);

            Assert.That(results.Select(r => r.Test), Is.EqualTo(new[]
            {
                "clsSampleTests.PassingTest", "clsSampleTests.FailingTest", "clsSampleTests.ErrorTest",
                "clsSampleTests.IgnoredTest", "clsSampleTests.RowTest.NotEqual"
            }));
        }

        [Test]
        public void Read_MapsStatusFromResultElement()
        {
            var results = JUnitResultReader.Read(JUnitXml);

            Assert.That(results.Select(r => r.Status), Is.EqualTo(new[]
            {
                TestStatus.Passed, TestStatus.Failed, TestStatus.Error, TestStatus.Ignored, TestStatus.Failed
            }));
        }

        [Test]
        public void Read_UsesMessageAttributeAndNullForEmptyMessage()
        {
            var results = JUnitResultReader.Read(JUnitXml);

            Assert.That(results[1].Message, Is.EqualTo("actual is less than expected (Expected: 2 but was: 1)"));
            Assert.That(results[0].Message, Is.Null);
            Assert.That(results[3].Message, Is.Null);
        }

        [Test]
        public void Read_ConvertsTimeInSecondsToMilliseconds()
        {
            var results = JUnitResultReader.Read(JUnitXml);

            Assert.That(results[0].ElapsedMilliseconds, Is.EqualTo(12).Within(0.0001));
        }
    }
}
