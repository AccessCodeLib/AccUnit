using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.AccUnit.Assertions.Interfaces;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;

namespace AccessCodeLib.AccUnit.TestRunner
{
    public class VbaTestRunner : ITestRunner
    {
        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event FinishedEventHandler TestSuiteFinished;
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

        public ITestResult Run(ITestSuite testSuite, ITestResultCollector testResultCollector)
        {
            RaiseTestSuiteStarted(testSuite);
            var results = new TestResultCollection(testSuite);

            foreach (var tests in testSuite.TestFixtures)
            {
                var result = Run(tests, testResultCollector);
                results.Add(result);
            }

            RaiseTestSuiteFinished(results);

            return results;
        }

        void RaiseTestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite, new TagList());
        }

        void RaiseTestSuiteFinished(ITestResult testResult)
        {
            TestSuiteFinished?.Invoke(testResult);
        }

        public ITestResult Run(ITestFixture testFixture, ITestResultCollector testResultCollector)
        {
            RaiseTestFixtureStarted(testFixture);

            var results = new TestResultCollection(testFixture);

            foreach (var test in testFixture.Tests)
            {
                var testResult = Run(test);
                testResultCollector.Add(testResult);
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

        public void Run(object testFixtureInstance, string testMethodName, ITestResultCollector testResultCollector = null)
        {
            var testFixture = new TestFixture(testFixtureInstance);

            if (testMethodName == "*")
            {
                testFixture.FillInstanceMembers(_vbProject);
                testFixture.FillTestListFromTestClassInstance(_vbProject);
                Run(testFixture, testResultCollector);
                return;
            }

            var test = CreateTest(testFixture, testMethodName);
            testFixture.Add(test);

            var result = Run(test);
            testResultCollector?.Add(result);
        }

        private ITest CreateTest(ITestFixture testFixture, string testMethodName)
        {
            var memberInfo = TestFixture.GetTestFixtureMember(_vbProject, testFixture.Name, testMethodName).TestClassMemberInfo;

            if (memberInfo.TestRows.Count > 0)
            {
                return new RowTest(testFixture, memberInfo);
            }

            var test = new MethodTest(testFixture, memberInfo);
            return test;
        }

        public ITestResult Run(IRowTest test)
        {
            var results = new TestResultCollection(test);
            foreach (var paramTest in test.ParamTests)
            {
                var result = Run(paramTest);
                results.Add(result);
            }
            return results;
        }

        public ITestResult Run(ITest test)
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
                return Run(rowTest);
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
                testResult.IsSuccess = true;
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
                    testResult.IsSuccess = false;
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

        void RaiseTestStarted(ITest test, IgnoreInfo ignoreInfo, TagList tags)
        {
            TestStarted?.Invoke(test, ignoreInfo, tags);
        }

        void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }
    }
}
