using AccessCodeLib.AccUnit.CodeCoverage;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("E4C80653-AD92-417F-AF25-9B084606FF13")]
    public interface ITestResultCollector
    {
        void Add(ITestResult TestResult);
    }

    public interface INotifyingTestResultCollector : ITestResultCollector, ITestResultCollectorEvents
    {
    }

    [ComVisible(true)]
    [Guid("95036E79-3476-4031-9656-E3762AEA5220")]
    public interface ITestResultSummaryPrinter
    {
        void PrintSummary(ITestSummary TestSummary, bool PrintTestResults = false);
    }

    public interface ITestSummaryTestResultCollector : ITestResultCollector
    {
        IEnumerable<ITestResult> TestResults { get; }
        ITestSummary Summary { get; }
    }

    public interface ITestResultCollectorEvents
    {
        event TestSuiteResetEventHandler TestSuiteReset;
        event TestSuiteStartedEventHandler TestSuiteStarted;
        event TestFixtureStartedEventHandler TestFixtureStarted;
        event TestStartedEventHandler TestStarted;
        event TestTraceMessageEventHandler TestTraceMessage;
        event FinishedEventHandler TestFinished;
        event TestResultEventHandler NewTestResult;
        event FinishedEventHandler TestFixtureFinished;
        event TestSuiteFinishedEventHandler TestSuiteFinished;
        event PrintSummaryEventHandler PrintSummary;
    }

    public delegate void TestResultEventHandler(ITestResult Result);
    public delegate void PrintSummaryEventHandler(ITestSummary TestSummary, bool PrintTestResults);
    
}
