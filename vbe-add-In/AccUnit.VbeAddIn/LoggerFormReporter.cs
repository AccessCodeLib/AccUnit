using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class LoggerFormReporter : ITestResultReporter
    {
        private LoggerForm _loggerForm;

        private INotifyingTestResultCollector _testResultCollector;

        public LoggerFormReporter()
        {
            LoggerForm.Visible = true;
        }

        private LoggerForm LoggerForm
        {
            get {
                if (_loggerForm == null)
                {
                    InitLoggerForm();
                }
                return _loggerForm;
            }
        }

        private void InitLoggerForm()
        {
            _loggerForm = new LoggerForm();
            _loggerForm.FormClosed += LoggerForm_FormClosed;
            _loggerForm.Visible = true; 
        }

        private void LoggerForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _loggerForm.FormClosed -= LoggerForm_FormClosed;
            _loggerForm = null;
        }

        public void Log(string message)
        {
            LogStringToTextBox(message);
        }   

        private void LogStringToTextBox(string message)
        {
            //append message to new line    
            LoggerForm.LogTextBox.AppendText(message + "\r\n");
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
            _testResultCollector.TestFixtureStarted += TestResultCollector_TestFixtureStarted;
            _testResultCollector.TestFixtureFinished += TestResultCollector_TestFixtureFinished;
            _testResultCollector.TestStarted += TestResultCollector_TestStarted;    
            _testResultCollector.TestFinished += TestResultCollector_TestFinished;  

        }

        private void TestResultCollector_TestFinished(ITestResult result)
        {
            //LogStringToTextBox("TestFinished");
        }

        private void TestResultCollector_TestStarted(ITest test, IgnoreInfo ignoreInfo)
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

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            LoggerForm.LogTextBox.Clear();
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
            LoggerForm.LogTextBox.Clear();
            LogStringToTextBox("TestSuite reset");
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            LogStringToTextBox(message);
        }
    }
}
