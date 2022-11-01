using AccessCodeLib.AccUnit.Interfaces;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("C856410C-BB3B-495E-8822-C8FBB4D4DC0F")]
    public interface IVBATestSuite : ITestSuite
    {
        #region COM visibility of inherited members

        new string Name { get; }
        ITestSummary Summary { get; }

        IVBATestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new IVBATestSuite Run();

        #endregion

        IVBATestSuite AddByClassName(string className);
        IVBATestSuite AddFromVBProject();

        [ComVisible(true)]
        void Dispose();

        [ComVisible(false)]
        TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername);

    }

    [ComVisible(true)]
    [Guid("5DAF3C2F-04AA-4B11-BB6B-6ADB15C7D554")]
    public interface IVBATestSuiteComInterface : IVBATestSuite
    {
        #region COM visibility of inherited members

        new ITestSummary Summary { get; }

        new IVBATestSuite Add([MarshalAs(UnmanagedType.IDispatch)] object testToAdd);
        new IVBATestSuite AddByClassName(string className);
        new IVBATestSuite AddFromVBProject();
        new IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData);
        new IVBATestSuite Run();
        new void Dispose();

        #endregion

        VBProject ActiveVBProject { get; set; }
        object HostApplication { [return: MarshalAs(UnmanagedType.IDispatch)] get; [param: MarshalAs(UnmanagedType.IDispatch)] set; }

    }

    [ComVisible(true)]
    [Guid("CE3393EA-8C3A-44E9-8191-8C35451E599E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestSuiteComEvents))]
    [ProgId("AccUnit.VBATestSuite")]
    public class VBATestSuite : AccUnit.VBATestSuite, IVBATestSuiteComInterface, IDisposable
    {
        
        IVBATestSuite IVBATestSuiteComInterface.Add(object testToAdd)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuite.Add(object testToAdd)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuiteComInterface.AddByClassName(string className)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuite.AddByClassName(string className)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuiteComInterface.AddFromVBProject()
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuite.AddFromVBProject()
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuiteComInterface.Reset(ResetMode mode)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuite.Reset(ResetMode mode)
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuiteComInterface.Run()
        {
            throw new NotImplementedException();
        }

        IVBATestSuite IVBATestSuite.Run()
        {
            throw new NotImplementedException();
        }
    }
}
