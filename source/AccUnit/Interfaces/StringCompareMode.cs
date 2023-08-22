using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("B7197C14-FBDB-4C2F-B5FE-2B535FB7558C")]
    [Flags]
    public enum StringCompareMode
    {
        BinaryCompare = 0,
        TextCompare = 1
    }
}


