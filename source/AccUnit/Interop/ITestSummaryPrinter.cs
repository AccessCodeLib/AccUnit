using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("4EE5E58E-C06A-489B-88D3-6801DD9843FB")]
    public interface ITestSummaryPrinter
    {
        void PrintSummary(bool PrintTestResults = false);
    }
}
