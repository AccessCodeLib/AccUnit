using AccessCodeLib.AccUnit.Configuration;
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

        public delegate void AccUnitFactoryRequestEventHandler(out IAccUnitFactory accUnitFactory);
        public event AccUnitFactoryRequestEventHandler AccUnitFactoryRequest;

        public delegate void CodeCoverageTrackerRequestEventHandler(out ICodeCoverageTracker codeCoverageTracker);
        public event CodeCoverageTrackerRequestEventHandler CodeCoverageTrackerRequest;

        public delegate void ErrorTrappingObserverRequestEventHandler(out IErrorTrappingObserver errorTrappingObserver);
        public event ErrorTrappingObserverRequestEventHandler ErrorTrappingObserverRequest;

        public delegate void HostApplicationInitializedEventHandler(object hostapplication);

        public event HostApplicationInitializedEventHandler HostApplicationInitialized;

        public object Application
        {
            set
            {
                HostApplicationInitialized?.Invoke(value);
            }
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object TestSuite(TestReportOutput OutputTo = TestReportOutput.DebugPrint)
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

        public IAccUnitFactory AccUnitFactory
        {
            get
            {
                IAccUnitFactory accUnitFactory = null;
                AccUnitFactoryRequest?.Invoke(out accUnitFactory);
                return accUnitFactory;
            }
        }

        public ICodeCoverageTracker CodeCoverageTracker
        {
            get
            {
                ICodeCoverageTracker codeCoverageTracker = null;
                CodeCoverageTrackerRequest?.Invoke(out codeCoverageTracker);
                return codeCoverageTracker;
            }
        }

        public IErrorTrappingObserver ErrorTrappingObserver
        {
            get
            {
                IErrorTrappingObserver errorTrappingObserver = null;
                ErrorTrappingObserverRequest?.Invoke(out errorTrappingObserver);
                return errorTrappingObserver;
            }
        }
    }

    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IAddInManagerBridge
    {
        object Application { set; }
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object TestSuite(TestReportOutput OutputTo = TestReportOutput.DebugPrint);
        IConstraintBuilder ConstraintBuilder { get; }
        IAssert Assert { get; }
        IAccUnitFactory AccUnitFactory { get; }
        ICodeCoverageTracker CodeCoverageTracker { get; }
        IErrorTrappingObserver ErrorTrappingObserver { get; }
    }

    public enum TestReportOutput
    {
        DebugPrint = 1,
        LogFile = 2
    }
}
