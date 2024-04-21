using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using System.Windows;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class LoggerControlReporter : ITestResultReporter
    {
        private readonly VbeUserControl<LoggerControl> _vbeUserControl;
        private readonly LoggerControl _loggerControl;

        private INotifyingTestResultCollector _testResultCollector;

        public LoggerControlReporter(VbeUserControl<LoggerControl> vbeUserControl)
        {
            _vbeUserControl = vbeUserControl;
            _loggerControl = vbeUserControl.Control;
        }

        private LoggerControl LoggerControl
        {
            get {
                return _loggerControl;
            }
        }

        public ITestResultCollector TestResultCollector
        {
            get { return _testResultCollector; }
            set
            {
                LogStringToTextBox("TestResultCollector set");
                _testResultCollector = value as INotifyingTestResultCollector;
                if (_testResultCollector == null)
                {
                    LogStringToTextBox("_testResultCollector is null");
                }
                InitEventHandler();
                LogStringToTextBox("InitEventHandler completed");
            }
        }

        private void InitEventHandler()
        {
            _testResultCollector.TestSuiteStarted += TestResultCollector_TestSuiteStarted;
            _testResultCollector.TestSuiteFinished += TestResultCollector_TestSuiteFinished;
            _testResultCollector.TestSuiteReset += TestResultCollector_TestSuiteReset;
            _testResultCollector.TestTraceMessage += TestResultCollector_TestTraceMessage;
            _testResultCollector.TestFixtureStarted += TestResultCollector_TestFixtureStarted;
            _testResultCollector.TestFixtureFinished += TestResultCollector_TestFixtureFinished;
            _testResultCollector.TestStarted += TestResultCollector_TestStarted;    
            _testResultCollector.TestFinished += TestResultCollector_TestFinished;  
        }

        private void TestResultCollector_TestFinished(ITestResult result)
        {
            //LogStringToTextBox("TestFinished");
        }

        private void TestResultCollector_TestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            LogStringToTextBox(test.DisplayName + "...");
        }

        private void TestResultCollector_TestFixtureFinished(ITestResult result)
        {
            //LogStringToTextBox("TestFixtureFinished");
        }

        private void TestResultCollector_TestFixtureStarted(ITestFixture fixture)
        {
            //LogStringToTextBox("TestFixtureStarted");
        }

        private void LogStringToTextBox(string message)
        {
            LoggerControl.LogTextBox.AppendText(message + "\r\n");
        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            LoggerControl.LogTextBox.Clear();
            _vbeUserControl.Show();
            LogStringToTextBox("TS started ...");
            if (testSuite is VBATestSuite vbaTestSuite)
                LogStringToTextBox(vbaTestSuite.ActiveVBProject.Name);
        }

        private void TestResultCollector_TestSuiteFinished(ITestSummary summary)
        {
            //LogStringToTextBox(summary.ToString());
            LogStringToTextBox("TS finished.");
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            LoggerControl.LogTextBox.Clear();
            LogStringToTextBox("TestSuite reset");
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            LogStringToTextBox(message);
        }
    }
}
