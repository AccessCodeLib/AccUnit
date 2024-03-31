using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("C51D526F-B545-4361-8DEB-147143C8550D")]
    [Flags]
    public enum VbaErrorTrapping : short
    {
        BreakOnAllErrors = 0,
        BreakInClassModule = 1,
        BreakOnUnhandledErrors = 2
    }
}
