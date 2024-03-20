using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("CC72AE5A-3C67-48BB-B8CE-C7D73506EC0A")]
    public interface IAccessTestSuite : Interfaces.IVBATestSuite
    {
        #region COM visibility of inherited members

        new string Name { get; }
        new object ActiveVBProject { [return: MarshalAs(UnmanagedType.IDispatch)] get; [param: MarshalAs(UnmanagedType.IDispatch)]  set; }
        new object HostApplication { [return: MarshalAs(UnmanagedType.IDispatch)] get; [param: MarshalAs(UnmanagedType.IDispatch)] set; }
        new ITestSummary Summary { get; }
        new ITestResultCollector TestResultCollector { get; set; }
        new ITestRunner TestRunner { get; set; }

        new IAccessTestSuite AppendTestResultReporter(ITestResultReporter reporter);
        new IAccessTestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IAccessTestSuite AddByClassName(string className);
        new IAccessTestSuite AddFromVBProject();
        new IAccessTestSuite Run();
        new IAccessTestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new void Dispose();

        #endregion

        IAccessTestSuite SelectTests(object TestNameFilter);
        IAccessTestSuite Filter(object FilterTags);
        ITestClassGenerator TestClassGenerator { get; }
    }

    [ComVisible(true)]
    [Guid("9F96EBE4-7FE4-4232-9510-A0818F9906FB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestSuiteComEvents))]
    [ProgId("AccUnit.AccessTestSuite")]
    public class AccessTestSuite : AccUnit.AccessTestSuite, IAccessTestSuite
    {
        ITestRunner IAccessTestSuite.TestRunner
        {
            get
            {
                return base.TestRunner as ITestRunner;
            }
            set
            {
                base.TestRunner = value;
            }
        }

        public new IAccessTestSuite Reset(ResetMode mode = ResetMode.ResetTestData)
        {
            base.Reset(mode);
            return this;
        }

        public new IAccessTestSuite Add(object testToAdd)
        {
            base.Add(testToAdd);
            return this;
        }

        public new IAccessTestSuite AddByClassName(string className)
        {
            base.AddByClassName(className);
            return this;
        }

        public new IAccessTestSuite AddFromVBProject()
        {
            base.AddFromVBProject();
            return this;
        }

        public new IAccessTestSuite Run()
        {
            base.Run();
            return this;
        }

        public IAccessTestSuite SelectTests(object TestNameFilter)
        {
            var testNameFilterEnumerable = InteropConverter.GetEnumerableFromFilterObject<string>(TestNameFilter);
            base.Select(testNameFilterEnumerable);
            return this;
        }

        public IAccessTestSuite Filter(object FilterTags)
        {
            IEnumerable<ITestItemTag> tags = InteropConverter.GetEnumerableFromFilterObject<ITestItemTag>(FilterTags);
            base.Filter(tags);
            return this;
        }

        public ITestClassGenerator TestClassGenerator
        {
            get
            {
                return new TestClassGenerator(ActiveVBProject);
            }
        }

        object IAccessTestSuite.ActiveVBProject
        {
            get { return base.ActiveVBProject; }
            set { base.ActiveVBProject = (VBProject)value;  }
        }

        public new IAccessTestSuite AppendTestResultReporter(ITestResultReporter reporter)
        {
            base.AppendTestResultReporter(reporter);
            return this;
        }

        protected override ITestResultCollector NewTestResultCollector()
        {
            return new TestResultCollector(this);
        }
    }
}
