using System;
using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools.Integration;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("F403650A-691E-427F-8E64-7729CD39C9E5")]
    public interface IVBATestSuite : Interfaces.IVBATestSuite
    {
        #region COM visibility of inherited members
        
        new string Name { get; }
        new VBProject ActiveVBProject { get; set; }
        new object HostApplication { [return: MarshalAs(UnmanagedType.IDispatch)] get; [param: MarshalAs(UnmanagedType.IDispatch)] set; }
        new ITestSummary Summary { get; }
        new ITestResultCollector TestResultCollector { get; set; }
        new ITestRunner TestRunner { get; set; }

        new IVBATestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IVBATestSuite AddByClassName(string className);
        new IVBATestSuite AddFromVBProject();
        new IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new IVBATestSuite Run();
        
        new void Dispose();

        #endregion

        ITestClassGenerator TestClassGenerator { get; }
    }

    [ComVisible(true)]
    [Guid("3824FB7F-768F-456E-8D43-5013628B8399")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestSuiteComEvents))]
    [ProgId("AccUnit.VBATestSuite")]
    public class VBATestSuite : AccUnit.VBATestSuite, IVBATestSuite, IDisposable
    {
        ITestRunner IVBATestSuite.TestRunner
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
        
        new public IVBATestSuite Run()
        {
            base.Run();
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
    }
}
