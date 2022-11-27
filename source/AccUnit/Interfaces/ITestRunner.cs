namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestRunner : ITestRunnerEvents
    {
        ITestResult Run(ITestSuite testSuite, ITestResultCollector testResultCollector);
        ITestResult Run(ITestFixture testFixture, ITestResultCollector testResultCollector);
        ITestResult Run(ITest test);
    }

    public interface ITestRunnerEvents
    {
        event TestSuiteStartedEventHandler TestSuiteStarted;
        event FinishedEventHandler TestSuiteFinished;
        event TestFixtureStartedEventHandler TestFixtureStarted;
        event FinishedEventHandler TestFixtureFinished;
        event TestStartedEventHandler TestStarted;
        event FinishedEventHandler TestFinished;
    }
}
