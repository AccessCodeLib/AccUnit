using AccessCodeLib.AccUnit.Configuration;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("B87911E4-A05D-4068-B456-411512B6BE78")]
    public interface IAccUnitFactory
    {
        IConstraintBuilder ConstraintBuilder { get; }
        IAssert Assert { get; }
        ITestRunner TestRunner([MarshalAs(UnmanagedType.IDispatch)] object VBProject);
        ITestBuilder TestBuilder { get; }
        IConfigurator Configurator([MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);
        IVBATestSuite VBATestSuite { get; }
        IAccessTestSuite AccessTestSuite { get; }
        ICodeCoverageTracker CodeCoverageTracker([MarshalAs(UnmanagedType.IDispatch)] object VBProject);
    }

    [ComVisible(true)]
    [Guid("79790592-4591-4004-A0E9-227ADD0E121F")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".AccUnitFactory")]
    public class AccUnitFactory : IAccUnitFactory
    {
        public IConstraintBuilder ConstraintBuilder
        {
            get
            {
                return new ConstraintBuilder();
            }
        }

        public IAssert Assert
        {
            get
            {
                return new Assert();
            }
        }

        public ITestRunner TestRunner(object vbProject = null)
        {
            return new TestRunner((VBProject)vbProject);
        }

        public ITestBuilder TestBuilder
        {
            get
            {
                return new TestBuilder();
            }
        }

        public IConfigurator Configurator(object vbProject = null)
        {
            return new Configurator((VBProject)vbProject);
        }

        public IVBATestSuite VBATestSuite
        {
            get
            {
                return new VBATestSuite();
            }
        }

        public IAccessTestSuite AccessTestSuite
        {
            get
            {
                return new AccessTestSuite();
            }
        }

        public ICodeCoverageTracker CodeCoverageTracker(object vbProject = null)
        {
            return new CodeCoverageTracker((VBProject)vbProject);
        }
    }
}
