using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.CodeCoverage;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("31B48F47-857E-4B65-8B45-4C4A13CD8E16")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestSuite : ITestSuiteEvents, ITestData
    {
        new string Name { get; } // Inherited from ITestData

        [ComVisible(false)]
        IEnumerable<ITestFixture> TestFixtures { get; }

        [ComVisible(false)]
        ITestSummary Summary { get; }

        [ComVisible(false)]
        ITestRunner TestRunner { get; }

        [ComVisible(false)]
        ITestResultCollector TestResultCollector { get; set; }

        [ComVisible(false)]
        ICodeCoverageTracker CodeCoverageTracker { get; set; }

        [ComVisible(false)]
        void AppendTestResultReporter(ITestResultReporter reporter);

        [ComVisible(false)]
        void AddTestClasses(IEnumerable<TestClassInfo> testClasses);

        [ComVisible(false)]
        ITestSuite Filter(IEnumerable<ITestItemTag> filterTags);

        [ComVisible(false)]
        ITestSuite Select(IEnumerable<string> methodFilter);

        ITestSuite Run();
        ITestSuite Reset(ResetMode mode = ResetMode.ResetTestData);

    }

    public interface ITestSuiteEvents
    {
        event TestSuiteStartedEventHandler TestSuiteStarted;
        event TestSuiteFinishedEventHandler TestSuiteFinished;
        event TestSuiteResetEventHandler TestSuiteReset;
        event TestFixtureStartedEventHandler TestFixtureStarted;
        event FinishedEventHandler TestFixtureFinished;
        event TestStartedEventHandler TestStarted;
        event FinishedEventHandler TestFinished;
        event TestTraceMessageEventHandler TestTraceMessage;
        event DisposeEventHandler Disposed;
    }

    public delegate void DisposeEventHandler(object sender);
    public delegate void NullReferenceEventHandler(ref object returnedObject);
    public delegate void FinishedEventHandler(ITestResult result);
    public delegate void TestSuiteStartedEventHandler(ITestSuite testSuite);
    public delegate void TestFixtureStartedEventHandler(ITestFixture fixture);
    public delegate void TestStartedEventHandler(ITest test, IgnoreInfo ignoreInfo);
    public delegate void MessageEventHandler(string message);
    public delegate void TestSuiteFinishedEventHandler(ITestSummary summary);
    public delegate void TestSuiteResetEventHandler(ResetMode resetmode, ref bool cancel);
    public delegate void TestTraceMessageEventHandler(string message, ICodeCoverageTracker CodeCoverageTracker);

}
