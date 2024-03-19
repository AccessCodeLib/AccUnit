using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("06F02DAF-5EBC-42F1-A2C4-8BEF443F54FA")]
    public interface ITestCollector : Interfaces.ITestResultCollector
    {
        new void Add(ITestResult testResult);
    }


    [ComVisible(true)]
    [Guid("14654A65-4377-44BA-961E-19DF332B18FD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestResultCollectorComEvents))]
    [ProgId("AccUnit.TestResultCollector")]
    public class TestResultCollector : Integration.TestResultCollector, ITestCollector
                                            , ITestResultCollectorEvents
    {
        public TestResultCollector(ITestData test) : base(test)
        {
        }

        public new void Add(ITestResult testResult)
        {
            base.Add(testResult);
        }

        protected override void RaiseTestSuiteStarted(ITestSuite testSuite)
        {
            TestSuiteStarted?.Invoke(testSuite);
        }

        protected override void RaiseTestTraceMessage(string message, CodeCoverage.ICodeCoverageTracker CodeCoverageTracker)
        {
            TestTraceMessage?.Invoke(message, CodeCoverageTracker as ICodeCoverageTracker);
        }

        protected override void RaiseNewTestResult(ITestResult testResult)
        {
            NewTestResult?.Invoke(testResult);
        }

        protected override void RaiseTestSuiteFinished(ITestSummary summary)
        {
            TestSuiteFinished?.Invoke(summary);
        }

        protected override void RaisePrintSummary(ITestSummary TestSummary, bool PrintTestResults)
        {
            PrintSummary?.Invoke(TestSummary, PrintTestResults);
        }

        public new event TestSuiteStartedEventHandler TestSuiteStarted;
        public new event TestTraceMessageEventHandler TestTraceMessage;
        public new event TestResultEventHandler NewTestResult;
        public new event TestSuiteFinishedEventHandler TestSuiteFinished;
        public new event PrintSummaryEventHandler PrintSummary;

    }

    public delegate void TestTraceMessageEventHandler(string Message, ICodeCoverageTracker CodeCoverageTracker);

}
