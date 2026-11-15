using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Builds the result of the MCP tool accunit_run_tests from the single test results.
    /// </summary>
    public static class TestResultMapper
    {
        public static TestRunResult CreateRunResult(IEnumerable<TestResultInfo> testResults,
                                                    IEnumerable<NotRunnableTestClass> notRunnableClasses,
                                                    bool includePassed)
        {
            var infos = testResults.ToList();
            var notRunnable = notRunnableClasses.ToList();

            var summary = new TestRunSummary
            {
                Total = infos.Count,
                Passed = infos.Count(i => i.Status == TestStatus.Passed),
                Failed = infos.Count(i => i.Status == TestStatus.Failed),
                Errors = infos.Count(i => i.Status == TestStatus.Error),
                Ignored = infos.Count(i => i.Status == TestStatus.Ignored),
                NotRunnable = notRunnable.Count,
                ElapsedMilliseconds = infos.Sum(i => i.ElapsedMilliseconds)
            };
            summary.Success = summary.Failed == 0 && summary.Errors == 0 && summary.NotRunnable == 0;

            return new TestRunResult
            {
                Summary = summary,
                Results = infos.Where(i => includePassed || i.Status != TestStatus.Passed).ToList(),
                NotRunnableClasses = notRunnable
            };
        }
    }
}
