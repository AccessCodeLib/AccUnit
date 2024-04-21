using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Integration
{
    public class TestResultCollector : TestResultCollection
                , INotifyingTestResultCollector, ITestResultCollector, ITestResultSummaryPrinter, ITestResultCollectorEvents
    {
        public TestResultCollector(ITestSuite testSuite) : base(testSuite)
        {
            TryInitTestSuiteListener(testSuite);
        }

        public new void Add(ITestResult testResult)
        {
            base.Add(testResult);
            RaiseNewTestResult(testResult);
        }

        private void TryInitTestSuiteListener(ITestSuite testSuite)
        {
            if (testSuite != null)
            {
                testSuite.TestSuiteReset += RaiseTestSuiteReset;    
                testSuite.TestSuiteStarted += RaiseTestSuiteStarted;
                testSuite.TestFixtureStarted += RaiseTestFixtureStarted;
                testSuite.TestStarted += RaiseTestStarted;
                testSuite.TestTraceMessage += RaiseTestTraceMessage;
                testSuite.TestFinished += RaiseTestFinished;
                testSuite.TestFixtureFinished += RaiseTestFixtureFinished;
                testSuite.TestSuiteFinished += RaiseTestSuiteFinished;
            }   
        }

        protected virtual void RaiseTestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            TestSuiteReset?.Invoke(resetmode, ref cancel);  
        }

        protected virtual void RaiseTestFixtureFinished(ITestResult result)
        {
            TestFixtureFinished?.Invoke(result);
        }

        protected virtual void RaiseTestFixtureStarted(ITestFixture fixture)
        {
            TestFixtureStarted?.Invoke(fixture);
        }

        protected virtual void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }

        protected virtual void RaiseTestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            //ignoreInfo.Ignore = false;
            TestStarted?.Invoke(test, ref ignoreInfo);
        }

        void ITestResultSummaryPrinter.PrintSummary(ITestSummary TestSummary, bool PrintTestResults)
        {
            RaisePrintSummary(TestSummary, PrintTestResults);
        }

        protected virtual void RaiseTestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite);
        }

        protected virtual void RaiseTestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            TestTraceMessage?.Invoke(message, CodeCoverageTracker); 
        }

        protected virtual void RaiseNewTestResult(ITestResult testResult)
        {
            NewTestResult?.Invoke(testResult);
        }

        protected virtual void RaiseTestSuiteFinished(ITestSummary summary)
        {
            TestSuiteFinished?.Invoke(summary);
        }

        protected virtual void RaisePrintSummary(ITestSummary TestSummary, bool PrintTestResults)
        {
            PrintSummary?.Invoke(TestSummary, PrintTestResults);
        }

        public event TestSuiteResetEventHandler TestSuiteReset;
        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event TestFixtureStartedEventHandler TestFixtureStarted;
        public event TestStartedEventHandler TestStarted;
        public event TestTraceMessageEventHandler TestTraceMessage;
        public event FinishedEventHandler TestFinished;
        public event TestResultEventHandler NewTestResult;
        public event FinishedEventHandler TestFixtureFinished;
        public event TestSuiteFinishedEventHandler TestSuiteFinished;
        public event PrintSummaryEventHandler PrintSummary;
    }

    
}