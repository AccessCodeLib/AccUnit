using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("291A478C-4878-4D73-9E2B-309CD3C5F908")]
    public interface ITestBuilder : IVBATestBuilder
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        new object CreateTest(string className);

        [return: MarshalAs(UnmanagedType.IDispatch)]
        new object CreateObject(string className);

        new void RefreshFactoryCodeModule();
        new void Dispose();
    }

    [ComVisible(true)]
    [Guid("E962986C-C46A-4DB2-A3C8-3B8623999B33")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TestBuilder")]
    public class TestBuilder : VBATestBuilder, ITestBuilder
    {
        public TestBuilder(IOfficeApplicationHelper applicationHelper)
                : base(applicationHelper)
        {
        }
    }
}
