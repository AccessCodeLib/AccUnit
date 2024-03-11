using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestRunner : ITestRunnerEvents
    {
        ITestResult Run(ITestSuite testSuite, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null);
        ITestResult Run(ITestFixture testFixture, ITestResultCollector testResultCollector, IEnumerable<string> methodFilter = null, IEnumerable<ITestItemTag> filterTags = null);
        ITestResult Run(ITest test, IEnumerable<ITestItemTag> filterTags);
    }

    public interface ITestRunnerEvents
    {
        event TestSuiteStartedEventHandler TestSuiteStarted;
        event TestSuiteFinishedEventHandler TestSuiteFinished;
        event TestFixtureStartedEventHandler TestFixtureStarted;
        event FinishedEventHandler TestFixtureFinished;
        event TestStartedEventHandler TestStarted;
        event FinishedEventHandler TestFinished;
    }
}
