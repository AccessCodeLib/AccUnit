using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestSuite : ITestSuiteEvents
    {
        string Name { get; }
        IEnumerable<ITestFixture> TestFixtures { get; }
        ITestSummary TestSummary { get; }

        ITestRunner TestRunner { get; set; }
        ITestSummaryTestResultCollector TestResultCollector { get; set; }
        
        ITestSuite Run();
        ITestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        void AddTestClasses(IEnumerable<TestClassInfo> testClasses);
    }
    
    public interface ITestSuiteEvents
    {
        event TestSuiteStartedEventHandler TestSuiteStarted;
        event FinishedEventHandler TestSuiteFinished;
        event TestSuiteResetEventHandler TestSuiteReset;
        event TestFixtureStartedEventHandler TestFixtureStarted;
        event FinishedEventHandler TestFixtureFinished;
        event TestStartedEventHandler TestStarted;
        event FinishedEventHandler TestFinished;
        event MessageEventHandler TestTraceMessage;
        event DisposeEventHandler Disposed;
    }
    
    public delegate void DisposeEventHandler(object sender);
    public delegate void NullReferenceEventHandler(ref object returnedObject);
    public delegate void FinishedEventHandler(ITestResult result);
    public delegate void TestSuiteStartedEventHandler(ITestSuite testSuite, TagList tags);
    public delegate void TestFixtureStartedEventHandler(ITestFixture fixture);
    public delegate void TestStartedEventHandler(ITest test, IgnoreInfo ignoreInfo, TagList tags);
    public delegate void MessageEventHandler(string message);
    public delegate void TestSuiteRunFinishedEventHandler(ITestSummary summary);
    public delegate void TestSuiteResetEventHandler(ResetMode resetmode, ref bool cancel);
}
