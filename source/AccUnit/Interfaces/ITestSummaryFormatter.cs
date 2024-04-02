namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestSummaryFormatter
    {
        int SeparatorMaxLength { get; set; }
        char SeparatorChar { get; set; }
        int TestFixtureFinishedSeparatorLength { get; set; }
        int TestCaseResultStartPos { get; set; }

        string GetTestSummaryText(ITestSummary summary);
        string GetTestCaseFinishedText(ITestResult result);
        string GetTestFixtureFinishedText(ITestResult result);
        string GetTestFixtureStartedText(ITestFixture fixture);
        string GetTestSuiteStartedText(ITestSuite suite);
        string GetTestSuiteFinishedText(ITestResult summary);
    }
}
