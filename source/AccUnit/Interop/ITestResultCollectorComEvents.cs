using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("95BC7D67-5088-4D6F-8835-94C5E0EA3738")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestResultCollectorComEvents : ITestResultCollectorEvents
    {
        new void TestSuiteStarted(ITestSuite TestSuite);
        new void TestTraceMessage(string Message, ICodeCoverageTracker CodeCoverageTracker);
        new void NewTestResult(ITestResult Result);
        new void TestSuiteFinished(ITestSummary Summary);
        new void PrintSummary(ITestSummary TestSummary, bool PrintTestResults);
    }
}
