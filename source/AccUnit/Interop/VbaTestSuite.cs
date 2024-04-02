﻿using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("F403650A-691E-427F-8E64-7729CD39C9E5")]
    public interface IVBATestSuite : Interfaces.IVBATestSuite
    {
        #region COM visibility of inherited members

        new string Name { get; }
        new ITestSummary Summary { get; }
        new ITestResultCollector TestResultCollector { get; set; }
        new ITestRunner TestRunner { get; }

        new IVBATestSuite AppendTestResultReporter(ITestResultReporter reporter);
        new IVBATestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IVBATestSuite AddByClassName(string className);
        new IVBATestSuite AddFromVBProject();
        new IVBATestSuite Run();
        new IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new void Dispose();

        #endregion

        IVBATestSuite SelectTests(object TestNameFilter);
        IVBATestSuite Filter(object FilterTags);
        ITestClassGenerator TestClassGenerator { get; }
        

    }

    [ComVisible(true)]
    [Guid("3824FB7F-768F-456E-8D43-5013628B8399")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestSuiteComEvents))]
    [ProgId("AccUnit.VBATestSuite")]
    public class VBATestSuite : AccUnit.VBATestSuite, IVBATestSuite, IDisposable
    {
        public VBATestSuite(OfficeApplicationHelper applicationHelper, IVBATestBuilder testBuilder, ITestRunner testRunner, ITestSummaryFormatter testSummaryFormatter)
               : base(applicationHelper, testBuilder, testRunner, testSummaryFormatter)
        {
        }

        ITestRunner IVBATestSuite.TestRunner
        {
            get
            {
                return base.TestRunner as ITestRunner;
            }
        }

        new public IVBATestSuite Add(object testToAdd)
        {
            base.Add(testToAdd);
            return this;
        }

        new public IVBATestSuite AddByClassName(string className)
        {
            base.AddByClassName(className);
            return this;
        }

        new public IVBATestSuite AddFromVBProject()
        {
            base.AddFromVBProject();
            return this;
        }

        new virtual public IVBATestSuite Reset(ResetMode mode)
        {
            base.Reset(mode);
            return this;
        }

        public new IVBATestSuite Run()
        {
            base.Run();
            return this;
        }

        public IVBATestSuite SelectTests(object TestNameFilter)
        {
            var testNameFilterEnumerable = InteropConverter.GetEnumerableFromFilterObject<string>(TestNameFilter);
            base.Select(testNameFilterEnumerable);
            return this;
        }

        public IVBATestSuite Filter(object FilterTags)
        {
            IEnumerable<ITestItemTag> tags = InteropConverter.GetEnumerableFromFilterObject<ITestItemTag>(FilterTags);
            base.Filter(tags);
            return this;
        }

        public ITestClassGenerator TestClassGenerator
        {
            get
            {
                /*
                var officeApplicationHelper = ComTools.GetTypeForComObject(HostApplication, "Access.Application") != null
                                                ? new AccessApplicationHelper(HostApplication) : new OfficeApplicationHelper(HostApplication);
                */
                return new TestClassGenerator(ActiveVBProject);
            }
        }

        public new IVBATestSuite AppendTestResultReporter(ITestResultReporter reporter)
        {
            base.AppendTestResultReporter(reporter);
            return this;
        }

        protected override ITestResultCollector NewTestResultCollector()
        {
            return new TestResultCollector(this);
        }

        //protected override void RaiseTraceMessage(string text)
        //{
        //    TestTraceMessage?.Invoke(text, CodeCoverageTracker as ICodeCoverageTracker);
        //}

        //public new event TestTraceMessageEventHandler TestTraceMessage;

    }
}
