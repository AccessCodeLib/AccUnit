using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.AccUnit.Assertions.Interfaces;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.TestRunner
{
    public class VbaTestRunner : ITestRunner
    {
        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event TestSuiteFinishedEventHandler TestSuiteFinished;
        public event TestFixtureStartedEventHandler TestFixtureStarted;
        public event FinishedEventHandler TestFixtureFinished;
        public event TestStartedEventHandler TestStarted;
        public event FinishedEventHandler TestFinished;

        private readonly VBProject _vbProject;

        public VbaTestRunner()
        {
        }

        public VbaTestRunner(VBProject vbProject)
        {
            _vbProject = vbProject;
        }

        public ITestResult Run(ITestSuite testSuite, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null)
        {
            RaiseTestSuiteStarted(testSuite);
            var results = new TestResultCollection(testSuite);

            foreach (var tests in testSuite.TestFixtures)
            {
                var result = Run(tests, testResultCollector, methodFilter, filterTags);
                results.Add(result);
            }

            RaiseTestSuiteFinished(results);

            return results;
        }

        void RaiseTestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite, new TagList());
        }

        void RaiseTestSuiteFinished(ITestSummary testSummary)
        {
            TestSuiteFinished?.Invoke(testSummary);
        }

        public ITestResult Run(ITestFixture testFixture, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null)
        {
            RaiseTestFixtureStarted(testFixture);

            var results = new TestResultCollection(testFixture);

            foreach (var test in testFixture.Tests)
            {
                if (methodFilter != null)
                {
                    if (methodFilter.Any(m => m.Contains("*") || m.Contains("?") || m.Contains("[")))
                    {
                        if (!PlaceholderFilterContainsTestName(methodFilter, test.Name))
                        {
                            continue;
                        }
                    }
                    else if (!methodFilter.Contains(test.Name))
                    {
                        continue;
                    }
                }

                var testResult = Run(test, filterTags);
                testResultCollector?.Add(testResult);
                results.Add(testResult);
            }

            RaiseTestFixtureFinished(results);

            return results;
        }

        private static bool PlaceholderFilterContainsTestName(IEnumerable<string> methodFilter, string testName)
        {
            var regExPattern = methodFilter.Aggregate("^", (current, filter) => current + filter.Replace("*", ".*").Replace("?", ".") + "|");
            regExPattern = regExPattern.Substring(0, regExPattern.Length - 1) + "$";
            return System.Text.RegularExpressions.Regex.IsMatch(testName, regExPattern);
        }

        void RaiseTestFixtureStarted(ITestFixture testFixture)
        {
            TestFixtureStarted?.Invoke(testFixture);
        }

        void RaiseTestFixtureFinished(ITestResult result)
        {
            TestFixtureFinished?.Invoke(result);
        }

        public ITestResult Run(object testFixtureInstance, string testMethodName, ITestResultCollector testResultCollector = null,
                               IEnumerable<ITestItemTag> filterTags = null)
        {
            var testFixture = new TestFixture(testFixtureInstance);

            if (filterTags != null && filterTags.Any())
            {
                testFixture.FillFixtureTags(_vbProject);
            }

            if (testMethodName == "*")
            {
                testFixture.FillInstanceMembers(_vbProject);
                testFixture.FillTestListFromTestClassInstance(_vbProject);
                return Run(testFixture, testResultCollector, null, filterTags);
            }

            var test = CreateTest(testFixture, testMethodName);
            testFixture.Add(test);

            var result = Run(test, filterTags);
            testResultCollector?.Add(result);

            return result;
        }

        private ITest CreateTest(ITestFixture testFixture, string testMethodName)
        {
            var memberInfo = TestFixture.GetTestFixtureMember(_vbProject, testFixture, testMethodName).TestClassMemberInfo;

            if (memberInfo.TestRows.Count > 0)
            {
                return new RowTest(testFixture, memberInfo);
            }

            var test = new MethodTest(testFixture, memberInfo);
            return test;
        }

        public ITestResult Run(IRowTest test, IEnumerable<ITestItemTag> filterTags = null)
        {
            var results = new TestResultCollection(test);
            foreach (var paramTest in test.ParamTests)
            {
                if (filterTags != null)
                {
                    if (!AllFilterTagsExists(paramTest.TestClassMemberInfo.Tags, filterTags))
                    {
                        continue;
                    }
                }

                var result = Run(paramTest);
                results.Add(result);
            }
            return results;
        }

        private static bool AllFilterTagsExists(IEnumerable<ITestItemTag> testTags, IEnumerable<ITestItemTag> filterTags)
        {
            return filterTags.All(tag => testTags.Any(testTag => testTag.Name.ToLower() == tag.Name.ToLower()));
        }

        public ITestResult Run(ITest test, IEnumerable<ITestItemTag> filterTags = null)
        {
            if (test.TestClassMemberInfo.IgnoreInfo.Ignore)
            {
                var ignoreTestResult = new TestResult(test)
                {
                    IsIgnored = true,
                    Message = test.TestClassMemberInfo.IgnoreInfo.Comment
                };
                RaiseTestFinished(ignoreTestResult);
                return ignoreTestResult;
            }

            if (test is IRowTest rowTest)
            {
                return Run(rowTest, filterTags);
            }

            if (filterTags != null && filterTags.Any())
            {
                if (!AllFilterTagsExists(test.TestClassMemberInfo.Tags, filterTags))
                {
                    var ignoreTestResult = new TestResult(test)
                    {
                        IsIgnored = true,
                        Message = "Test ignored because of tag filter"
                    };
                    RaiseTestFinished(ignoreTestResult);
                    return ignoreTestResult;
                }
            }

            var testResult = new TestResult(test);
            var testFixture = test.Fixture;

            using (var invocationHelper = new InvocationHelper(testFixture.Instance))
            {
                var ignoreInfo = new IgnoreInfo();

                RaiseTestStarted(test, ignoreInfo, test.TestClassMemberInfo.Tags);
                if (ignoreInfo.Ignore)
                {
                    testResult.IsIgnored = true;
                    testResult.Message = ignoreInfo.Comment;
                    RaiseTestFinished(testResult);
                    return testResult;
                }

                RunTestSetup(testFixture, invocationHelper);
                RunTest(test, testResult, invocationHelper);
                AssertThrowsStore.CompaireTestRunnerException(null, testResult);
                RunTeardown(testFixture, invocationHelper);
            }

            RaiseTestFinished(testResult);
            return testResult;
        }

        private static void RunTestSetup(ITestFixture testFixture, InvocationHelper invocationHelper)
        {
            if (testFixture.HasSetup)
            {
                invocationHelper.InvokeMethod(testFixture.Members.Setup.Name);
            }
        }

        private static void RunTest(ITest test, TestResult testResult, InvocationHelper invocationHelper)
        {
            try
            {
                if (test is IParamTest paramTest)
                {
                    var testParams = paramTest.Parameters.ToArray();
                    invocationHelper.InvokeMethod(test.MethodName, testParams);
                }
                else
                {
                    invocationHelper.InvokeMethod(test.MethodName);
                }
                testResult.IsPassed = true;
            }
            catch (Exception ex)
            {
                Exception messageException;
                bool IsInvocationException = false;
                if (ex is System.Reflection.TargetInvocationException)
                {
                    messageException = ex.InnerException ?? ex;
                    IsInvocationException = true;
                }
                else
                {
                    messageException = ex;
                }

                if (!AssertThrowsStore.CompaireTestRunnerException(messageException, testResult))
                {
                    if (messageException is AssertionException assertionException)
                    {
                        testResult.IsFailure = true;
                        testResult.Message = assertionException.Message;
                    }
                    else
                    {
                        // AssertionException does not seem to get through via Invoke

                        if (IsInvocationException && messageException.GetType() == typeof(Exception))
                        {
                            testResult.IsFailure = true;
                        }
                        else
                        {
                            testResult.IsError = true;
                        }
                        testResult.Message += messageException.Message;
                    }
                    testResult.IsPassed = false;
                }
            }
            finally
            {
                testResult.Executed = true;
            }
        }

        private static void RunTeardown(ITestFixture testFixture, InvocationHelper invocationHelper)
        {
            if (testFixture.HasTeardown)
            {
                invocationHelper.InvokeMethod(testFixture.Members.Teardown.Name);
            }
        }

        void RaiseTestStarted(ITest test, IgnoreInfo ignoreInfo, ITagList tags)
        {
            TestStarted?.Invoke(test, ignoreInfo, tags);
        }

        void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }
    }
}
