using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace AccessCodeLib.AccUnit.Integration
{
    /// <summary>
    /// Writes a test summary in the JUnit XML format used by CI systems
    /// (one testsuite per test class, one testcase per test or test row).
    /// </summary>
    public class JUnitXmlWriter
    {
        private const string TestSuitesName = "AccUnit";
        private const string TimestampFormat = "yyyy-MM-ddTHH:mm:ss";
        private const string TimeFormat = "0.000";
        private const double MillisecondsPerSecond = 1000.0;
        private const string NotExecutedMessage = "Not executed";

        public void WriteToFile(ITestSummary summary, string outputFilePath, DateTime timestamp)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is missing.", nameof(outputFilePath));

            var fullPath = Path.GetFullPath(outputFilePath);
            var directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            using (var stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                Write(summary, stream, timestamp);
            }
        }

        public void Write(ITestSummary summary, Stream output, DateTime timestamp)
        {
            if (summary == null)
                throw new ArgumentNullException(nameof(summary));
            if (output == null)
                throw new ArgumentNullException(nameof(output));

            var suites = GroupTestCasesBySuite(summary.TestResults);
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = true
            };

            using (var writer = XmlWriter.Create(output, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("testsuites");
                writer.WriteAttributeString("name", TestSuitesName);
                WriteCounts(writer, suites.SelectMany(s => s.TestCases).ToList());
                writer.WriteAttributeString("timestamp", FormatTimestamp(timestamp));

                foreach (var suite in suites)
                {
                    WriteTestSuite(writer, suite, timestamp);
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static List<JUnitTestSuite> GroupTestCasesBySuite(IEnumerable<ITestResult> testResults)
        {
            var suites = new List<JUnitTestSuite>();
            var suitesByName = new Dictionary<string, JUnitTestSuite>(StringComparer.OrdinalIgnoreCase);

            foreach (var testCase in testResults.SelectMany(GetTestCaseResults))
            {
                var suiteName = GetSuiteName(testCase.Test);
                if (!suitesByName.TryGetValue(suiteName, out JUnitTestSuite suite))
                {
                    suite = new JUnitTestSuite(suiteName);
                    suitesByName.Add(suiteName, suite);
                    suites.Add(suite);
                }
                suite.TestCases.Add(testCase);
            }

            return suites;
        }

        // A row test result holds one result per row; each row becomes its own testcase.
        private static IEnumerable<ITestResult> GetTestCaseResults(ITestResult result)
        {
            if (result is ITestSummary rowResults)
                return rowResults.TestResults.SelectMany(GetTestCaseResults);

            return new[] { result };
        }

        private static string GetSuiteName(ITestData test)
        {
            if (test is ITest methodTest && methodTest.Fixture != null)
                return methodTest.Fixture.Name;

            return test?.Name ?? string.Empty;
        }

        private static string GetTestCaseName(ITestData test, string suiteName)
        {
            if (test == null)
                return string.Empty;

            var fullName = test.FullName ?? string.Empty;
            var suitePrefix = suiteName + ".";
            if (suiteName.Length > 0 && fullName.StartsWith(suitePrefix, StringComparison.OrdinalIgnoreCase))
                return fullName.Substring(suitePrefix.Length);

            return test.Name ?? fullName;
        }

        private static void WriteTestSuite(XmlWriter writer, JUnitTestSuite suite, DateTime timestamp)
        {
            writer.WriteStartElement("testsuite");
            writer.WriteAttributeString("name", ToXmlText(suite.Name));
            WriteCounts(writer, suite.TestCases);
            writer.WriteAttributeString("timestamp", FormatTimestamp(timestamp));

            foreach (var testCase in suite.TestCases)
            {
                WriteTestCase(writer, testCase, suite.Name);
            }

            writer.WriteEndElement();
        }

        private static void WriteCounts(XmlWriter writer, IList<ITestResult> testCases)
        {
            writer.WriteAttributeString("tests", FormatCount(testCases.Count));
            writer.WriteAttributeString("failures", FormatCount(testCases.Count(r => r.IsFailure)));
            writer.WriteAttributeString("errors", FormatCount(testCases.Count(r => r.IsError && !r.IsFailure)));
            writer.WriteAttributeString("skipped", FormatCount(testCases.Count(IsSkipped)));
            writer.WriteAttributeString("time", FormatTime(testCases.Sum(r => r.ElapsedTime)));
        }

        private static void WriteTestCase(XmlWriter writer, ITestResult result, string suiteName)
        {
            writer.WriteStartElement("testcase");
            writer.WriteAttributeString("name", ToXmlText(GetTestCaseName(result.Test, suiteName)));
            writer.WriteAttributeString("classname", ToXmlText(suiteName));
            writer.WriteAttributeString("time", FormatTime(result.ElapsedTime));

            if (result.IsFailure)
                WriteResultElement(writer, "failure", result.Message);
            else if (result.IsError)
                WriteResultElement(writer, "error", result.Message);
            else if (result.IsIgnored)
                WriteResultElement(writer, "skipped", result.Message);
            else if (!result.IsPassed)
                WriteResultElement(writer, "skipped", NotExecutedMessage);

            writer.WriteEndElement();
        }

        private static bool IsSkipped(ITestResult result)
        {
            return !result.IsFailure && !result.IsError && (result.IsIgnored || !result.IsPassed);
        }

        private static void WriteResultElement(XmlWriter writer, string elementName, string message)
        {
            var text = ToXmlText(message);
            writer.WriteStartElement(elementName);
            writer.WriteAttributeString("message", text);
            if (text.Length > 0)
                writer.WriteString(text);
            writer.WriteEndElement();
        }

        // Messages from VBA can contain characters that are not allowed in XML (e.g. vbNullChar).
        private static string ToXmlText(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            var builder = new StringBuilder(value.Length);
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (char.IsHighSurrogate(c) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    builder.Append(c).Append(value[i + 1]);
                    i++;
                }
                else if (XmlConvert.IsXmlChar(c))
                {
                    builder.Append(c);
                }
            }
            return builder.ToString().Trim();
        }

        private static string FormatCount(int count)
        {
            return count.ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatTime(double elapsedMilliseconds)
        {
            return (elapsedMilliseconds / MillisecondsPerSecond).ToString(TimeFormat, CultureInfo.InvariantCulture);
        }

        private static string FormatTimestamp(DateTime timestamp)
        {
            return timestamp.ToString(TimestampFormat, CultureInfo.InvariantCulture);
        }

        private class JUnitTestSuite
        {
            public JUnitTestSuite(string name)
            {
                Name = name;
            }

            public string Name { get; }
            public List<ITestResult> TestCases { get; } = new List<ITestResult>();
        }
    }
}
