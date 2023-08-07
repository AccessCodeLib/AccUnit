using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("997CA0A7-ADBA-464C-9C0E-D0DCA85BED1E")]
    [Flags]
    public enum RunState
    {
        Runnable = 0,
        NotRunnable = 1,
        Ignored = 2,
        Executed = 3
    }
}
