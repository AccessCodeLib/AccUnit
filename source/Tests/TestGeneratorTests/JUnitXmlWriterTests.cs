using AccessCodeLib.AccUnit;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace TestGeneratorTests
{
    public class JUnitXmlWriterTests
    {
        private static readonly DateTime Timestamp = new DateTime(2026, 11, 12, 14, 30, 5);

        [Test]
        public void Write_TestsOfTwoClasses_WritesOneTestSuitePerClass()
        {
            var summary = NewSummary(
                PassedResult("ClassA", "Test1"),
                PassedResult("ClassB", "Test1"),
                PassedResult("ClassA", "Test2"));

            var document = WriteToDocument(summary);

            var suites = document.Root.Elements("testsuite").ToList();
            Assert.That(suites.Select(s => (string)s.Attribute("name")), Is.EqualTo(new[] { "ClassA", "ClassB" }));
            Assert.That((string)suites[0].Attribute("tests"), Is.EqualTo("2"));
            Assert.That((string)suites[1].Attribute("tests"), Is.EqualTo("1"));
        }

        [Test]
        public void Write_PassedTest_WritesTestCaseWithoutResultElement()
        {
            var summary = NewSummary(PassedResult("ClassA", "Test1"));

            var testCase = WriteToDocument(summary).Descendants("testcase").Single();

            Assert.That((string)testCase.Attribute("name"), Is.EqualTo("Test1"));
            Assert.That((string)testCase.Attribute("classname"), Is.EqualTo("ClassA"));
            Assert.That(testCase.HasElements, Is.False);
        }

        [Test]
        public void Write_FailedTest_WritesFailureWithMessage()
        {
            var result = NewResult("ClassA", "Test1");
            result.IsFailure = true;
            result.Message = "Expected: 1 but was: 2";

            var failure = WriteToDocument(NewSummary(result)).Descendants("testcase").Single().Element("failure");

            Assert.That(failure, Is.Not.Null);
            Assert.That((string)failure.Attribute("message"), Is.EqualTo("Expected: 1 but was: 2"));
            Assert.That(failure.Value, Is.EqualTo("Expected: 1 but was: 2"));
        }

        [Test]
        public void Write_TestWithError_WritesErrorWithMessage()
        {
            var result = NewResult("ClassA", "Test1");
            result.IsError = true;
            result.Message = "SetUp: Division by zero";

            var error = WriteToDocument(NewSummary(result)).Descendants("testcase").Single().Element("error");

            Assert.That(error, Is.Not.Null);
            Assert.That((string)error.Attribute("message"), Is.EqualTo("SetUp: Division by zero"));
        }

        [Test]
        public void Write_IgnoredTest_WritesSkippedWithReason()
        {
            var result = NewResult("ClassA", "Test1");
            result.IsIgnored = true;
            result.Message = "not implemented yet";

            var skipped = WriteToDocument(NewSummary(result)).Descendants("testcase").Single().Element("skipped");

            Assert.That(skipped, Is.Not.Null);
            Assert.That((string)skipped.Attribute("message"), Is.EqualTo("not implemented yet"));
        }

        [Test]
        public void Write_RowTest_WritesOneTestCasePerRow()
        {
            var fixture = new FakeTestFixture("ClassA");
            var rowTest = new FakeTest(fixture, "RowTest1", "ClassA.RowTest1");
            var rowResults = new TestResultCollection(rowTest);
            rowResults.Add(new TestResult(new FakeTest(fixture, "RowTest1", "ClassA.RowTest1.1")) { IsPassed = true });
            var failedRow = new TestResult(new FakeTest(fixture, "RowTest1", "ClassA.RowTest1.2")) { IsFailure = true, Message = "row failed" };
            rowResults.Add(failedRow);

            var testCases = WriteToDocument(NewSummary(rowResults)).Descendants("testcase").ToList();

            Assert.That(testCases.Select(t => (string)t.Attribute("name")), Is.EqualTo(new[] { "RowTest1.1", "RowTest1.2" }));
            Assert.That(testCases[1].Element("failure"), Is.Not.Null);
        }

        [Test]
        public void Write_MixedResults_WritesCountsOnTestSuites()
        {
            var failed = NewResult("ClassA", "Failed");
            failed.IsFailure = true;
            var error = NewResult("ClassA", "Error");
            error.IsError = true;
            var ignored = NewResult("ClassB", "Ignored");
            ignored.IsIgnored = true;

            var root = WriteToDocument(NewSummary(PassedResult("ClassA", "Passed"), failed, error, ignored)).Root;

            Assert.That(root.Name.LocalName, Is.EqualTo("testsuites"));
            Assert.That((string)root.Attribute("tests"), Is.EqualTo("4"));
            Assert.That((string)root.Attribute("failures"), Is.EqualTo("1"));
            Assert.That((string)root.Attribute("errors"), Is.EqualTo("1"));
            Assert.That((string)root.Attribute("skipped"), Is.EqualTo("1"));
        }

        [Test]
        public void Write_GermanCulture_WritesTimeInSecondsWithDecimalPoint()
        {
            var result = PassedResult("ClassA", "Test1");
            result.ElapsedTime = 1234.5;
            var summary = NewSummary(result);

            XDocument document;
            var currentCulture = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-AT");
                document = WriteToDocument(summary);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }

            Assert.That((string)document.Descendants("testcase").Single().Attribute("time"), Is.EqualTo("1.235"));
            Assert.That((string)document.Root.Attribute("timestamp"), Is.EqualTo("2026-11-12T14:30:05"));
        }

        [Test]
        public void Write_MessageWithCharactersNotAllowedInXml_RemovesThoseCharacters()
        {
            var result = NewResult("ClassA", "Test1");
            result.IsFailure = true;
            result.Message = "a\0b <c> & \"d\"";

            var failure = WriteToDocument(NewSummary(result)).Descendants("failure").Single();

            Assert.That((string)failure.Attribute("message"), Is.EqualTo("ab <c> & \"d\""));
        }

        [Test]
        public void WriteToFile_DirectoryDoesNotExist_CreatesDirectoryAndFile()
        {
            var directory = Path.Combine(Path.GetTempPath(), "AccUnitJUnitXmlWriterTests", Guid.NewGuid().ToString("N"));
            var filePath = Path.Combine(directory, "results", "junit.xml");
            try
            {
                new JUnitXmlWriter().WriteToFile(NewSummary(PassedResult("ClassA", "Test1")), filePath, Timestamp);

                Assert.That(XDocument.Load(filePath).Descendants("testcase").Count(), Is.EqualTo(1));
            }
            finally
            {
                if (Directory.Exists(directory))
                    Directory.Delete(directory, true);
            }
        }

        [Test]
        public void Reporter_TestSuiteFinished_WritesFile()
        {
            var filePath = Path.Combine(Path.GetTempPath(), "AccUnitJUnitXmlReporterTests_" + Guid.NewGuid().ToString("N") + ".xml");
            var collector = new RaisingTestResultCollector();
            var reporter = new JUnitXmlTestResultReporter(filePath) { TestResultCollector = collector };
            try
            {
                collector.RaiseFinished(NewSummary(PassedResult("ClassA", "Test1")));

                Assert.That(XDocument.Load(filePath).Descendants("testcase").Count(), Is.EqualTo(1));
            }
            finally
            {
                reporter.TestResultCollector = null;
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
        }

        private static XDocument WriteToDocument(ITestSummary summary)
        {
            using (var stream = new MemoryStream())
            {
                new JUnitXmlWriter().Write(summary, stream, Timestamp);
                stream.Position = 0;
                return XDocument.Load(stream);
            }
        }

        private static ITestSummary NewSummary(params ITestResult[] results)
        {
            var summary = new TestResultCollection(null);
            foreach (var result in results)
            {
                summary.Add(result);
            }
            return summary;
        }

        private static TestResult NewResult(string fixtureName, string methodName)
        {
            var test = new FakeTest(new FakeTestFixture(fixtureName), methodName, fixtureName + "." + methodName);
            return new TestResult(test) { Executed = true };
        }

        private static TestResult PassedResult(string fixtureName, string methodName)
        {
            var result = NewResult(fixtureName, methodName);
            result.IsPassed = true;
            return result;
        }

        private class RaisingTestResultCollector : TestResultCollector
        {
            public RaisingTestResultCollector() : base(null)
            {
            }

            public void RaiseFinished(ITestSummary summary)
            {
                RaiseTestSuiteFinished(summary);
            }
        }

        private class FakeTestFixture : ITestFixture
        {
            public FakeTestFixture(string name)
            {
                Name = name;
            }

            public string Name { get; }
            public string FullName { get { return Name; } }
            public object Parent { get { return null; } }
            public object Instance { get { return null; } }
            public RunState RunState { get; set; }
            public ITestFixtureMembers Members { get { return null; } }
            public IEnumerable<ITest> Tests { get { return Enumerable.Empty<ITest>(); } }
            public IEnumerable<ITestItemTag> Tags { get { return Enumerable.Empty<ITestItemTag>(); } }
            public bool HasFixtureSetup { get; set; }
            public bool HasSetup { get; set; }
            public bool HasTeardown { get; set; }
            public bool HasFixtureTeardown { get; set; }
        }

        private class FakeTest : ITest
        {
            public FakeTest(ITestFixture fixture, string methodName, string fullName)
            {
                Fixture = fixture;
                MethodName = methodName;
                FullName = fullName;
            }

            public string Name { get { return MethodName; } }
            public string FullName { get; }
            public object Parent { get { return Fixture; } }
            public ITestFixture Fixture { get; }
            public string MethodName { get; }
            public string DisplayName { get; set; }
            public RunState RunState { get; set; }
            public ITestClassMemberInfo TestClassMemberInfo { get { return null; } }
        }
    }
}
