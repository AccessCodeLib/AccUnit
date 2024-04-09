using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class TestResultReporter : ITestResultReporter
    {
        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event TestSuiteFinishedEventHandler TestSuiteFinished;
        public event TestSuiteResetEventHandler TestSuiteReset;
        public event TestFixtureStartedEventHandler TestFixtureStarted;
        public event FinishedEventHandler TestFixtureFinished;
        public event TestStartedEventHandler TestStarted;
        public event FinishedEventHandler TestFinished;
        public event TestTraceMessageEventHandler TestTraceMessage;
        
        private INotifyingTestResultCollector _testResultCollector;  

        public ITestResultCollector TestResultCollector {
            get { return _testResultCollector; }
            set {
                _testResultCollector = value as INotifyingTestResultCollector;
                InitEventHandler();
            }   
        }

        private void InitEventHandler()
        {
            _testResultCollector.TestSuiteStarted += TestResultCollector_TestSuiteStarted;  
            _testResultCollector.TestSuiteFinished += TestResultCollector_TestSuiteFinished;
            _testResultCollector.TestSuiteReset += TestResultCollector_TestSuiteReset;

            _testResultCollector.TestFixtureStarted += TestResultCollector_TestFixtureStarted;
            _testResultCollector.TestFixtureFinished += TestResultCollector_TestFixtureFinished;

            _testResultCollector.TestStarted += TestResultCollector_TestStarted;    
            _testResultCollector.TestFinished += TestResultCollector_TestFinished;  
            _testResultCollector.TestTraceMessage += TestResultCollector_TestTraceMessage;
            
        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite);
        }

        private void TestResultCollector_TestSuiteFinished(ITestSummary summary)
        {
            TestSuiteFinished?.Invoke(summary);   
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            TestSuiteReset?.Invoke(resetmode, ref cancel);
        }

        private void TestResultCollector_TestFixtureStarted(ITestFixture fixture)
        {
            TestFixtureStarted?.Invoke(fixture);    
        }

        private void TestResultCollector_TestFixtureFinished(ITestResult result)
        {
            TestFixtureFinished?.Invoke(result);    
        }

        private void TestResultCollector_TestStarted(ITest test, IgnoreInfo ignoreInfo)
        {
            TestStarted?.Invoke(test, ignoreInfo);  
        }

        private void TestResultCollector_TestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);   
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            TestTraceMessage?.Invoke(message, CodeCoverageTracker); 
        }
    }
}
