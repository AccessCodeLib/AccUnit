﻿using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("570D48B4-989D-47CD-852F-F6F8AFE6DD14")]
    public interface ITestRunner : Interfaces.ITestRunner
    {
        ITestResult Run([MarshalAs(UnmanagedType.IDispatch)] object TestFixtureInstance,
                        string TestMethodName = "*",
                        ITestResultCollector TestResultCollector = null,
                        object filterTags = null);
    }

    [ComVisible(true)]
    [Guid("DBED9DB2-5F34-46A4-87B1-7CB3C4FB94F5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TestRunner")]
    public class TestRunner : AccUnit.TestRunner.VbaTestRunner, ITestRunner
    {
        public TestRunner(VBProject vbProject = null) : base(vbProject)
        {
        }

        public ITestResult Run([MarshalAs(UnmanagedType.IDispatch)] object TestFixtureInstance, string TestMethodName = "*",
                                ITestResultCollector TestResultCollector = null,
                                object FilterTags = null)
        {

            IEnumerable<ITestItemTag> tags = FilterTags != null ? InteropConverter.GetEnumerableFromFilterObject<ITestItemTag>(FilterTags) : null;
            return base.Run(TestFixtureInstance, TestMethodName, TestResultCollector, tags);
        }
    }
}
