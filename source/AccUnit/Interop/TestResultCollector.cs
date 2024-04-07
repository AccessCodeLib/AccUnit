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
                                            , INotifyingTestResultCollector, ITestResultCollectorEvents
                                            , ITestResultSummaryPrinter
    {
        public TestResultCollector(ITestSuite test) : base(test)
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

        protected override void RaiseTestStarted(ITest test, IgnoreInfo ignoreInfo)
        {
            TestStarted?.Invoke(test, ignoreInfo);
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

        protected override void RaiseTestFixtureStarted(ITestFixture testFixture)
        {
            TestFixtureStarted?.Invoke(testFixture);
        }

        protected override void RaiseTestFixtureFinished(ITestResult result)
        {
           TestFixtureFinished?.Invoke(result);
        }   

        protected override void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }

        protected override void RaiseTestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            TestSuiteReset?.Invoke(resetmode, ref cancel);
        }   

        public new event TestSuiteResetEventHandler TestSuiteReset;
        public new event TestSuiteStartedEventHandler TestSuiteStarted;
        public new event TestFixtureStartedEventHandler TestFixtureStarted;
        public new event TestStartedEventHandler TestStarted;
        public new event TestTraceMessageEventHandler TestTraceMessage;
        public new event FinishedEventHandler TestFinished;
        public new event TestResultEventHandler NewTestResult;
        public new event FinishedEventHandler TestFixtureFinished;
        public new event TestSuiteFinishedEventHandler TestSuiteFinished;
        public new event PrintSummaryEventHandler PrintSummary;
    }

    //public delegate void TestSuiteStartedEventHandler(ITestSuite testSuite);
    public delegate void TestTraceMessageEventHandler(string Message, ICodeCoverageTracker CodeCoverageTracker);
    //public delegate void TestStartedEventHandler(ITest test, IgnoreInfo ignoreInfo);

}
