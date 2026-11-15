using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.McpServer
{
    public class TestRunResult
    {
        public TestRunSummary Summary { get; set; }

        /// <summary>Failed, error and ignored tests; passed tests only if requested.</summary>
        public List<TestResultInfo> Results { get; set; } = new List<TestResultInfo>();

        /// <summary>Test classes that could not be prepared for a run from the MCP server.</summary>
        public List<NotRunnableTestClass> NotRunnableClasses { get; set; } = new List<NotRunnableTestClass>();

        /// <summary>Changes the server made to the database to be able to run the tests.</summary>
        public List<string> EnvironmentChanges { get; set; } = new List<string>();
    }

    public class TestRunSummary
    {
        public int Total { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Errors { get; set; }
        public int Ignored { get; set; }
        public int NotRunnable { get; set; }
        public double ElapsedMilliseconds { get; set; }
        public bool Success { get; set; }
    }

    public class TestResultInfo
    {
        public string Test { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
        public double ElapsedMilliseconds { get; set; }
    }

    public class NotRunnableTestClass
    {
        public string ClassName { get; set; }
        public string Reason { get; set; }
    }

    public static class TestStatus
    {
        public const string Passed = "passed";
        public const string Failed = "failed";
        public const string Error = "error";
        public const string Ignored = "ignored";
        public const string NotExecuted = "notExecuted";
    }
}
