﻿using Microsoft.Vbe.Interop;
using System;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface IVBATestSuite : ITestSuite, IDisposable
    {
        IVBATestSuite Add(object testToAdd);
        IVBATestSuite AddByClassName(string className);
        IVBATestSuite AddFromVBProject();
        TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername);

        VBProject ActiveVBProject { get; set; }
        object HostApplication { get; set; }
    }
}
