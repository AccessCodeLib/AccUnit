using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("E111F33A-7F56-400C-8D6E-5807EF06F29B")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITestSummary
    {
        [ComVisible(false)]
        IEnumerable<ITestResult> TestResults { get; }

        double ElapsedTime { get; }
        int Total { get; }
        int Passed { get; }
        int Failed { get; }
        int Error { get; }
        int Ignored { get; }
        void Reset();
    }
}
