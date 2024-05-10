using AccessCodeLib.AccUnit.Interop;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    [ComVisible(true)]
    [Guid("E69DF056-1CB6-4977-8554-78F7FFF6BA0A")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.AddInManager")]
    public class AddInManagerBridge : IAddInManagerBridge
    {
        public delegate void TestSuiteRequestEventHandler(out Interfaces.IVBATestSuite testsuite);
        public event TestSuiteRequestEventHandler TestSuiteRequest;

        public delegate void ConstraintBuilderRequestEventHandler(out IConstraintBuilder constraintBuilder);
        public event ConstraintBuilderRequestEventHandler ConstraintBuilderRequest;

        public delegate void AssertRequestEventHandler(out IAssert assert);
        public event AssertRequestEventHandler AssertRequest;

        public delegate void HostApplicationInitializedEventHandler(object hostapplication);

        public event HostApplicationInitializedEventHandler HostApplicationInitialized;

        public object Application
        {
            set
            {
                HostApplicationInitialized?.Invoke(value);
            }
        }

        public Interfaces.IVBATestSuite TestSuite(TestReportOutput OutputTo = TestReportOutput.DebugPrint)
        {
            Interfaces.IVBATestSuite suite = null;
            TestSuiteRequest?.Invoke(out suite);
            return suite;
        }

        public IConstraintBuilder ConstraintBuilder
        {
            get
            {
                IConstraintBuilder constraintBuilder = null;
                ConstraintBuilderRequest?.Invoke(out constraintBuilder);
                return constraintBuilder;
            }
        }

        public IAssert Assert
        {
            get
            {
                IAssert assert = null;
                AssertRequest?.Invoke(out assert);
                return assert;
            }
        }
    }

    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IAddInManagerBridge
    {
        object Application { set; }
        Interfaces.IVBATestSuite TestSuite(TestReportOutput OutputTo = TestReportOutput.DebugPrint);
        IConstraintBuilder ConstraintBuilder { get; }
        IAssert Assert { get; }
    }

    public enum TestReportOutput
    {
        DebugPrint = 1,
        LogFile = 2
    }
}
