using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Integration
{
    /// <summary>
    /// Writes the test results as JUnit XML file when the test suite has finished.
    /// </summary>
    [ComVisible(true)]
    [Guid("6D1F2B7E-3C4A-4E8B-9A57-2F0C8D61B4E3")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Interop.Constants.ProgIdLibName + ".JUnitXmlTestResultReporter")]
    public class JUnitXmlTestResultReporter : ITestResultReporter
    {
        private readonly JUnitXmlWriter _writer = new JUnitXmlWriter();
        private ITestResultCollector _testResultCollector;
        private DateTime _testSuiteStartTime = DateTime.Now;

        public JUnitXmlTestResultReporter(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is missing.", nameof(outputFilePath));

            OutputFilePath = outputFilePath;
        }

        public string OutputFilePath { get; }

        public ITestResultCollector TestResultCollector
        {
            get { return _testResultCollector; }
            set
            {
                RemoveEventHandlers(_testResultCollector as ITestResultCollectorEvents);
                _testResultCollector = value;
                AddEventHandlers(_testResultCollector as ITestResultCollectorEvents);
            }
        }

        private void AddEventHandlers(ITestResultCollectorEvents collectorEvents)
        {
            if (collectorEvents == null)
                return;

            collectorEvents.TestSuiteStarted += OnTestSuiteStarted;
            collectorEvents.TestSuiteFinished += OnTestSuiteFinished;
        }

        private void RemoveEventHandlers(ITestResultCollectorEvents collectorEvents)
        {
            if (collectorEvents == null)
                return;

            collectorEvents.TestSuiteStarted -= OnTestSuiteStarted;
            collectorEvents.TestSuiteFinished -= OnTestSuiteFinished;
        }

        private void OnTestSuiteStarted(ITestSuite testSuite)
        {
            _testSuiteStartTime = DateTime.Now;
        }

        private void OnTestSuiteFinished(ITestSummary summary)
        {
            _writer.WriteToFile(summary, OutputFilePath, _testSuiteStartTime);
        }
    }
}
