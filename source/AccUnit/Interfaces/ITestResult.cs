using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("A45A8733-3069-40C0-A896-80C94EBDA1D8")]
    public interface ITestResult
    {
        ITestData Test { get; }

        bool Executed { get; }
        bool IsError { get; }
        bool IsFailure { get; }
        bool IsIgnored { get; }
        bool IsSuccess { get; }
        string Message { get; }

        string Result { get; }
        double ElapsedTime { set; get; }
    }
}
