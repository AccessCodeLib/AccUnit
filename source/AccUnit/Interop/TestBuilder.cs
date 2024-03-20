using Microsoft.Vbe.Interop;
using System;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("291A478C-4878-4D73-9E2B-309CD3C5F908")]
    public interface ITestBuilder
    {
        object HostApplication
        {
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
            [param: MarshalAs(UnmanagedType.IDispatch)]
            set;
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object CreateTest(string className);

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object CreateObject(string className);

        object ActiveVBProject { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        void RefreshFactoryCodeModule();
        void Dispose();
    }

    [ComVisible(true)]
    [Guid("E962986C-C46A-4DB2-A3C8-3B8623999B33")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TestBuilder")]
    public class TestBuilder : VBATestBuilder, ITestBuilder
    {
        object ITestBuilder.ActiveVBProject => (object)base.ActiveVBProject;
    }
}
