using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    [ComVisible(true)]
    [Guid("1ED5A466-959A-4679-B29C-8B2A5EA7E5F4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgIdAttribute("AccUnit.AddInManager")]
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
                if (HostApplicationInitialized != null)
                {
                    HostApplicationInitialized(value);
                }
            }
        }

        public IVBATestSuite TestSuite
        {
            get
            {
                IVBATestSuite suite;
                TestSuiteRequest(out suite);
                return suite;
            }
        }

    }

    [ComVisible(true)]
    [Guid("81C905A1-E218-4743-BA0F-D1DB0033ABF9")]
    public interface IAddInManagerBridge
    {
        IVBATestSuite TestSuite { get; }
        object Application { set; }
    }
}
