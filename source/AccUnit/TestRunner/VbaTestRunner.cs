using AccessCodeLib.AccUnit.Assertions.Interfaces;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;

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


        public void Run(ITestSuite testSuite, ITestResultCollector testResultCollector)
        {
            RaiseTestSuiteStarted(testSuite);
            throw new NotImplementedException();
            RaiseTestSuiteFinished(null);
        }

        void RaiseTestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite, new TagList());
        }

        void RaiseTestSuiteFinished(ITestResult testResult)
        {
            TestSuiteFinished?.Invoke(testResult);
        }

        public void Run(ITestFixture testFixture, ITestResultCollector testResultCollector)
        {
            RaiseTestFixtureStarted(testFixture);
            
            var result = new TestResultCollection(testFixture);

            foreach (var test in testFixture.Tests)
            {
                var testResult = Run(test);
                testResultCollector.Add(testResult);
                result.Add(testResult);
            }
            
            RaiseTestFixtureFinished(result);
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
                Run(testFixture, testResultCollector);
                return;
            }

            var test = new MethodTest(testFixture, testMethodName);
            testFixture.Add(test);

            var result = Run(test);
            if (testResultCollector != null)
            {
                testResultCollector.Add(result);
            }
        }

        public ITestResult Run(ITest test)
        {
            var testResult = new TestResult(test);
            var testFixture = test.Fixture;

            using (var invocationHelper = new InvocationHelper(testFixture.Instance))
            {
                var ignoreInfo = new IgnoreInfo();

                RaiseTestStarted(test, ignoreInfo, new TagList());
                if (ignoreInfo.Ignore)
                {
                    testResult.IsIgnored = true;
                    testResult.Message = ignoreInfo.Comment;
                    RaiseTestFinished(testResult);
                    return testResult;
                }

                if (testFixture.HasSetup)
                {
                    invocationHelper.InvokeMethod(testFixture.Members.Setup.Name);
                }
                
                try
                {
                    invocationHelper.InvokeMethod(test.MethodName);
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

                    if (messageException is AssertionException assertionException)
                    {
                        testResult.IsFailure = true;
                        testResult.Message = assertionException.Message;
                    }
                    else
                    {
                        // über Invoke kommt AssertionException anscheinend nicht durch
                        
                        if (IsInvocationException && messageException.GetType() == typeof(Exception))
                        {
                            testResult.IsFailure = true;
                        }
                        else
                        {
                            testResult.IsError = true;
                        }
                        testResult.Message = messageException.Message;
                    }
                    testResult.IsSuccess = false;
                }
                finally
                {
                    testResult.Executed = true;
                }
                
                if (testFixture.HasTeardown)
                {
                    invocationHelper.InvokeMethod(testFixture.Members.Teardown.Name);
                }
            }

            RaiseTestFinished(testResult);
            return testResult;
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
