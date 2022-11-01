using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("5A88F46E-015B-4F0B-93D3-394AA9FB5B6F")]
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
        double Time { set; get; }
    }
}
