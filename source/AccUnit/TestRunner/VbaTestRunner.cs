using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.AccUnit.Assertions.Interfaces;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        public VbaTestRunner(VBProject vbProject)
        {
            _vbProject = vbProject;
        }

        public ITestResult Run(ITestSuite testSuite, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null)
        {
            RaiseTestSuiteStarted(testSuite);
            var results = new TestResultCollection(testSuite);

            foreach (var fixture in testSuite.TestFixtures)
            {
                var result = Run(fixture, testResultCollector, methodFilter, filterTags);
                results.Add(result);
            }

            RaiseTestSuiteFinished(results);

            return results;
        }

        void RaiseTestSuiteStarted(ITestSuite testSuite/*, IEnumerable<ITestItemTag> tags*/)
        {
            TestSuiteStarted?.Invoke(testSuite);
        }

        void RaiseTestSuiteFinished(ITestSummary testSummary)
        {
            TestSuiteFinished?.Invoke(testSummary);
        }

        public ITestResult Run(ITestFixture testFixture, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null)
        {
            RaiseTestFixtureStarted(testFixture);

            var results = new TestResultCollection(testFixture);

            TestNameFilterMatcher placeholderMatcher = null;
            if (methodFilter != null && methodFilter.Any(TestNameFilterMatcher.ContainsWildcard))
            {
                placeholderMatcher = new TestNameFilterMatcher(methodFilter);
            }

            foreach (var test in testFixture.Tests)
            {
                if (methodFilter != null)
                {
                    if (placeholderMatcher != null)
                    {
                        if (!placeholderMatcher.IsMatch(test.Name))
                        {
                            continue;
                        }
                    }
                    else if (!methodFilter.Contains(test.Name, StringComparer.OrdinalIgnoreCase))
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

            // single test run needs the fixture members as well (Setup/Teardown)
            testFixture.FillInstanceMembers(_vbProject);

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
            var ignoreInfo = new IgnoreInfo();
            RaiseTestStarted(test, ref ignoreInfo);

            if (ignoreInfo.Ignore)
            {
                var ignoreTestResult = new TestResult(test)
                {
                    IsIgnored = true,
                    Message = ignoreInfo.Comment
                };
                RaiseTestFinished(ignoreTestResult);
                return ignoreTestResult;
            }

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

            RaiseTestFinished(results);
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
                RaiseTestStarted(test, ref ignoreInfo);
                if (ignoreInfo.Ignore)
                {
                    testResult.IsIgnored = true;
                    testResult.Message = ignoreInfo.Comment;
                    RaiseTestFinished(testResult);
                    return testResult;
                }

                var stopwatch = Stopwatch.StartNew();
                if (TryRunTestSetup(testFixture, testResult, invocationHelper))
                {
                    RunTest(test, testResult, invocationHelper);
                    AssertThrowsStore.CompaireTestRunnerException(null, testResult);
                }
                TryRunTeardown(testFixture, testResult, invocationHelper);
                testResult.ElapsedTime = stopwatch.Elapsed.TotalMilliseconds;
            }

            RaiseTestFinished(testResult);
            return testResult;
        }

        private static Exception GetMessageException(Exception ex)
        {
            return ex is System.Reflection.TargetInvocationException ? ex.InnerException ?? ex : ex;
        }

        // Setup error = test error (like NUnit); the test itself is not executed.
        private static bool TryRunTestSetup(ITestFixture testFixture, TestResult testResult, InvocationHelper invocationHelper)
        {
            if (!testFixture.HasSetup)
            {
                return true;
            }

            try
            {
                invocationHelper.InvokeMethod(testFixture.Members.Setup.Name);
                return true;
            }
            catch (Exception ex)
            {
                // do not let an expected error number leak into the next test
                AssertThrowsStore.Clear();

                testResult.IsPassed = false;
                testResult.IsError = true;
                testResult.Executed = true;
                testResult.Message = "SetUp: " + GetMessageException(ex).Message;
                return false;
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

        // Teardown error is reported additionally; an already determined failure/error is kept.
        private static void TryRunTeardown(ITestFixture testFixture, TestResult testResult, InvocationHelper invocationHelper)
        {
            if (!testFixture.HasTeardown)
            {
                return;
            }

            try
            {
                invocationHelper.InvokeMethod(testFixture.Members.Teardown.Name);
            }
            catch (Exception ex)
            {
                var message = "TearDown: " + GetMessageException(ex).Message;
                if (testResult.IsFailure || testResult.IsError)
                {
                    testResult.Message = string.IsNullOrEmpty(testResult.Message)
                        ? message
                        : testResult.Message + Environment.NewLine + message;
                }
                else
                {
                    testResult.IsPassed = false;
                    testResult.IsError = true;
                    testResult.Message = message;
                }
                testResult.Executed = true;
            }
        }

        void RaiseTestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            TestStarted?.Invoke(test, ref ignoreInfo);
        }

        void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }
    }
}
