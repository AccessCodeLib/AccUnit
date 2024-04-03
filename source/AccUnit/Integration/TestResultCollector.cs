using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Integration
{
    public class TestResultCollector : TestResultCollection
                , ITestResultCollector, ITestResultSummaryPrinter, ITestResultCollectorEvents
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
                testSuite.TestSuiteStarted += RaiseTestSuiteStarted;
                testSuite.TestTraceMessage += RaiseTestTraceMessage;
                testSuite.TestSuiteFinished += RaiseTestSuiteFinished;
            }   
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

        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event TestTraceMessageEventHandler TestTraceMessage;
        public event TestResultEventHandler NewTestResult;
        public event TestSuiteFinishedEventHandler TestSuiteFinished;
        public event PrintSummaryEventHandler PrintSummary;
    }

    
}