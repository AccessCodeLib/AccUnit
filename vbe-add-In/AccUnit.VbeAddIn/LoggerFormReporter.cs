using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class LoggerFormReporter : ITestResultReporter
    {
        private readonly Form _loggerForm;   
        private readonly TextBox _logTextBox;

        private INotifyingTestResultCollector _testResultCollector;

        public LoggerFormReporter(LoggerForm loggerForm)
        {
            _loggerForm = loggerForm;   
            _logTextBox = loggerForm.LogTextBox;
        }

        public ITestResultCollector TestResultCollector
        {
            get { return _testResultCollector; }
            set
            {
                _testResultCollector = value as INotifyingTestResultCollector;
                InitEventHandler();
            }
        }

        private void InitEventHandler()
        {
            _testResultCollector.TestSuiteStarted += TestResultCollector_TestSuiteStarted;
            _testResultCollector.TestSuiteFinished += TestResultCollector_TestSuiteFinished;
            _testResultCollector.TestSuiteReset += TestResultCollector_TestSuiteReset;
            _testResultCollector.TestTraceMessage += TestResultCollector_TestTraceMessage;

        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            _logTextBox.Clear();
            _logTextBox.AppendText("TS started ...");
        }

        private void TestResultCollector_TestSuiteFinished(ITestSummary summary)
        {
            _logTextBox.AppendText(summary.ToString());
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            _logTextBox.Clear();
            _logTextBox.AppendText("TestSuite reset");
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            _logTextBox.AppendText(message);
        }

    }
}
