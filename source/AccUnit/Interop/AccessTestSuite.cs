using System;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("CC72AE5A-3C67-48BB-B8CE-C7D73506EC0A")]
    public interface IAccessTestSuite : Interfaces.IVBATestSuite
    {
        #region COM visibility of inherited members

        new string Name { get; }
        new VBProject ActiveVBProject { get; set; }
        new object HostApplication { [return: MarshalAs(UnmanagedType.IDispatch)] get; [param: MarshalAs(UnmanagedType.IDispatch)] set; }
        new ITestSummary Summary { get; }
        new ITestResultCollector TestResultCollector { get; set; }
        new ITestRunner TestRunner { get; set; }

        new IAccessTestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IAccessTestSuite AddByClassName(string className);
        new IAccessTestSuite AddFromVBProject();
        new IAccessTestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new IAccessTestSuite Run();

        new void Dispose();

        #endregion

        ITestClassGenerator TestClassGenerator { get; }
    }
    
    [ComVisible(true)]
    [Guid("9F96EBE4-7FE4-4232-9510-A0818F9906FB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestSuiteComEvents))]
    [ProgIdAttribute("AccUnit.AccessTestSuite")]
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

        public ITestClassGenerator TestClassGenerator
        {
            get
            {
                return new TestClassGenerator(ActiveVBProject);
            }
        }
    }
    
}
