using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    [ComVisible(true)]
    [Guid("E69DF056-1CB6-4977-8554-78F7FFF6BA0A")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.AddInManager")]
    public class AddInManagerBridge : IAddInManagerBridge
    {
        public delegate void TestSuiteRequestEventHandler(out IVBATestSuite testsuite);
        public event TestSuiteRequestEventHandler TestSuiteRequest;

        public delegate void HostApplicationInitializedEventHandler(object hostapplication);

        public event HostApplicationInitializedEventHandler HostApplicationInitialized;

        public object Application 
        { 
            set
            {
                HostApplicationInitialized?.Invoke(value);
            }
        }

        public IVBATestSuite TestSuite
        {
            get
            {
                TestSuiteRequest(out IVBATestSuite suite);
                return suite;
            }
        }
    }

    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IAddInManagerBridge
    {
        IVBATestSuite TestSuite { get; }
        object Application { set; }
    }
}
