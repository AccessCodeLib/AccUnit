using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("E4C80653-AD92-417F-AF25-9B084606FF13")]
    public interface ITestResultCollector
    {
        void Add(ITestResult testResult);
    }

    public interface ITestSummaryTestResultCollector : ITestResultCollector
    {
        IEnumerable<ITestResult> TestResults { get; }
        ITestSummary Summary { get; }
    }
}
