using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("C80C791F-7C12-4CFA-AD63-DBF428BFA10D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestSuiteComEvents : ITestSuiteEvents
    {
        new void TestTraceMessage(string Message, ICodeCoverageTracker CodeCoverageTracker);
        new void TestSuiteStarted(ITestSuite TestSuite);
    }
}
