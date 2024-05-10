using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class LoggerBoxControlReporter : ITestResultReporter, INotifyPropertyChanged
    {
        private readonly VbeUserControl<LoggerBoxControl> _vbeUserControl;
        private readonly LoggerBoxControl _loggerControl;

        private INotifyingTestResultCollector _testResultCollector;

        public event PropertyChangedEventHandler PropertyChanged;

        public LoggerBoxControlReporter(VbeUserControl<LoggerBoxControl> vbeUserControl)
        {
            _vbeUserControl = vbeUserControl;
            _loggerControl = vbeUserControl.Control;
        }

        private void OnPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public ITestResultCollector TestResultCollector
        {
            get { return _testResultCollector; }
            set
            {
                _testResultCollector = value as INotifyingTestResultCollector;
                if (_testResultCollector == null)
                {
                    LogStringToTextBox("TestResultCollector is null");
                }
                InitEventHandler();
            }
        }

        public ObservableCollection<string> LogMessages { get; } = new ObservableCollection<string>();

        public string LogMessagesText
        {
            get { return string.Join("\r\n", LogMessages); }
        }

        private void LogStringToTextBox(string message)
        {
            _loggerControl.LoggerTextBox.AppendText(message + "\r\n");
        }

        private void ClearLogMessages()
        {
            LogMessages.Clear();
            _loggerControl.LoggerTextBox.Clear();
            OnPropertyChanged(nameof(LogMessagesText));
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
            //LogStringToTextBox(test.DisplayName + "...");
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
            ClearLogMessages();
            _vbeUserControl.Show();
        }

        private void TestResultCollector_TestSuiteFinished(ITestSummary summary)
        {
            //
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            ClearLogMessages();
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            LogStringToTextBox(message);
        }
    }

    public static class LoggerBoxControlInfo
    {
        public const string ProgID = @"AccUnit.VbeAddIn.LoggerBoxControlInfo";
        public const string PositionGuid = @"3DE8AA0C-8D8D-427F-B8BD-14594B60BCE1";
    }
}
