﻿using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [Guid("D6C5FF83-10A3-4C88-80DF-29068F252B5F")]
    public interface ITestData
    {
        string Name { get; }
        string FullName { get; }
    }
}
