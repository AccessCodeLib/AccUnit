using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Reads the test results from the JUnit XML file written by AccUnit (JUnitXmlTestResultReporter).
    /// </summary>
    public static class JUnitResultReader
    {
        private const double MillisecondsPerSecond = 1000;

        public static List<TestResultInfo> Read(string junitXml)
        {
            return XDocument.Parse(junitXml)
                .Descendants("testcase")
                .Select(ToTestResultInfo)
                .ToList();
        }

        private static TestResultInfo ToTestResultInfo(XElement testCase)
        {
            var resultElement = testCase.Elements().FirstOrDefault(e => e.Name == "failure" || e.Name == "error" || e.Name == "skipped");
            return new TestResultInfo
            {
                Test = GetTestName(testCase),
                Status = GetStatus(resultElement),
                Message = GetMessage(resultElement),
                ElapsedMilliseconds = GetElapsedMilliseconds(testCase)
            };
        }

        private static string GetTestName(XElement testCase)
        {
            var className = (string)testCase.Attribute("classname");
            var name = (string)testCase.Attribute("name");
            return string.IsNullOrEmpty(className) ? name : className + "." + name;
        }

        private static string GetStatus(XElement resultElement)
        {
            if (resultElement == null) return TestStatus.Passed;
            switch (resultElement.Name.LocalName)
            {
                case "failure": return TestStatus.Failed;
                case "error": return TestStatus.Error;
                default: return TestStatus.Ignored;
            }
        }

        private static string GetMessage(XElement resultElement)
        {
            var message = (string)resultElement?.Attribute("message");
            return string.IsNullOrEmpty(message) ? null : message;
        }

        private static double GetElapsedMilliseconds(XElement testCase)
        {
            double seconds;
            return double.TryParse((string)testCase.Attribute("time"), NumberStyles.Float, CultureInfo.InvariantCulture, out seconds)
                ? seconds * MillisecondsPerSecond
                : 0;
        }
    }
}
