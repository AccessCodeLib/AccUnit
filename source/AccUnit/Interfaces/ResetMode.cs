using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("DAEE570F-DD0F-4C31-AEA5-64875C557FD0")]
    [Flags]
    public enum ResetMode
    {
        None = 0,
        ResetTestData = 1,
        RemoveTests = 2,
        ResetTestSuite = 4,
        RefreshFactoryModule = 8,
        DeleteFactoryModule = 16
    }
}


