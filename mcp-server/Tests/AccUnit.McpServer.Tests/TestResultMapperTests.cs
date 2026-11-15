using System.Linq;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class TestResultMapperTests
    {
        [Test]
        public void CreateRunResult_CountsResultsAndListsOnlyNotPassed()
        {
            var results = new[]
            {
                Result("clsTests.PassingTest", TestStatus.Passed, 1),
                Result("clsTests.FailingTest", TestStatus.Failed, 2),
                Result("clsTests.ErrorTest", TestStatus.Error),
                Result("clsTests.IgnoredTest", TestStatus.Ignored),
                Result("clsTests.RowTest.Equal", TestStatus.Passed)
            };

            var runResult = TestResultMapper.CreateRunResult(results, new NotRunnableTestClass[0], includePassed: false);

            Assert.That(runResult.Summary.Total, Is.EqualTo(5));
            Assert.That(runResult.Summary.Passed, Is.EqualTo(2));
            Assert.That(runResult.Summary.Failed, Is.EqualTo(1));
            Assert.That(runResult.Summary.Errors, Is.EqualTo(1));
            Assert.That(runResult.Summary.Ignored, Is.EqualTo(1));
            Assert.That(runResult.Summary.ElapsedMilliseconds, Is.EqualTo(3));
            Assert.That(runResult.Summary.Success, Is.False);
            Assert.That(runResult.Results.Select(r => r.Test),
                        Is.EquivalentTo(new[] { "clsTests.FailingTest", "clsTests.ErrorTest", "clsTests.IgnoredTest" }));
        }

        [Test]
        public void CreateRunResult_IncludePassed_ListsAllResults()
        {
            var results = new[] { Result("a", TestStatus.Passed), Result("b", TestStatus.Failed) };

            var runResult = TestResultMapper.CreateRunResult(results, new NotRunnableTestClass[0], includePassed: true);

            Assert.That(runResult.Results.Count, Is.EqualTo(2));
        }

        [Test]
        public void CreateRunResult_OnlyPassedAndIgnored_Success()
        {
            var results = new[] { Result("a", TestStatus.Passed), Result("b", TestStatus.Ignored) };

            var runResult = TestResultMapper.CreateRunResult(results, new NotRunnableTestClass[0], includePassed: false);

            Assert.That(runResult.Summary.Success, Is.True);
        }

        [Test]
        public void CreateRunResult_NotRunnableClass_NoSuccess()
        {
            var notRunnable = new[] { new NotRunnableTestClass { ClassName = "clsX", Reason = "r" } };

            var runResult = TestResultMapper.CreateRunResult(new[] { Result("a", TestStatus.Passed) }, notRunnable, includePassed: false);

            Assert.That(runResult.Summary.NotRunnable, Is.EqualTo(1));
            Assert.That(runResult.Summary.Success, Is.False);
            Assert.That(runResult.NotRunnableClasses.Single().ClassName, Is.EqualTo("clsX"));
        }

        private static TestResultInfo Result(string test, string status, double elapsed = 0)
        {
            return new TestResultInfo { Test = test, Status = status, ElapsedMilliseconds = elapsed };
        }
    }
}
