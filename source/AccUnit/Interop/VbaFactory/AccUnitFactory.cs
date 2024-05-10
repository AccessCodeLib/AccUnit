using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Integration;
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
        ITestBuilder TestBuilder([MarshalAs(UnmanagedType.IDispatch)] object hostApplication);
        IConfigurator Configurator([MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);
        IVBATestSuite VBATestSuite([MarshalAs(UnmanagedType.IDispatch)] object hostApplication, ITestBuilder testBuilder = null, ITestRunner testRunner = null, ITestSummaryFormatter testSummaryFormatter = null, ITestResultCollector externalTestResultCollector = null);
        IAccessTestSuite AccessTestSuite([MarshalAs(UnmanagedType.IDispatch)] object hostApplication, ITestBuilder testBuilder = null, ITestRunner testRunner = null, ITestSummaryFormatter testSummaryFormatter = null, ITestResultCollector externalTestResultCollector = null);
        ICodeCoverageTracker CodeCoverageTracker([MarshalAs(UnmanagedType.IDispatch)] object VBProject);
        IErrorTrappingObserver AccessErrorTrappingObserver([MarshalAs(UnmanagedType.IDispatch)] object HostApplication);

        [ComVisible(false)]
        IVBATestSuite VBATestSuite(IOfficeApplicationHelper applicationHelper, ITestBuilder testBuilder = null, ITestRunner testRunner = null, ITestSummaryFormatter testSummaryFormatter = null, ITestResultCollector externalTestResultCollector = null);
        [ComVisible(false)]
        IAccessTestSuite AccessTestSuite(IAccessApplicationHelper applicationHelper, ITestBuilder testBuilder = null, ITestRunner testRunner = null, ITestSummaryFormatter testSummaryFormatter = null, ITestResultCollector externalTestResultCollector = null);

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

        public ITestBuilder TestBuilder(object hostApplication)
        {
            return new TestBuilder(GetApplicationHelper(hostApplication));
        }

        public IConfigurator Configurator(object vbProject = null)
        {
            return new Configurator((VBProject)vbProject);
        }

        public IVBATestSuite VBATestSuite(
                                    object hostApplication,
                                    ITestBuilder testBuilder = null,
                                    ITestRunner testRunner = null,
                                    ITestSummaryFormatter testSummaryFormatter = null,
                                    ITestResultCollector externalTestResultCollector = null)
        {
            var applicationHelper = GetApplicationHelper(hostApplication);
            return VBATestSuite(applicationHelper, testBuilder, testRunner, testSummaryFormatter, externalTestResultCollector);
        }

        public IVBATestSuite VBATestSuite(
                                    IOfficeApplicationHelper applicationHelper,
                                    ITestBuilder testBuilder = null,
                                    ITestRunner testRunner = null,
                                    ITestSummaryFormatter testSummaryFormatter = null,
                                    ITestResultCollector externalTestResultCollector = null)
        {
            if (testRunner == null)
                testRunner = new TestRunner(applicationHelper.CurrentVBProject);

            if (testBuilder == null)
                testBuilder = new TestBuilder(applicationHelper);

            if (testSummaryFormatter == null)
                testSummaryFormatter = new TestSummaryFormatter(TestSuiteUserSettings.Current.SeparatorMaxLength, TestSuiteUserSettings.Current.SeparatorChar);

            var testSuite = new VBATestSuite(applicationHelper, testBuilder, testRunner, testSummaryFormatter);

            if (externalTestResultCollector != null)
                testSuite.TestResultCollector = externalTestResultCollector;

            return testSuite;
        }

        private OfficeApplicationHelper GetApplicationHelper(object hostApplication)
        {
            return ComTools.GetTypeForComObject(hostApplication, "Access.Application") != null
                ? new AccessApplicationHelper(hostApplication) : new OfficeApplicationHelper(hostApplication);
        }

        public IAccessTestSuite AccessTestSuite(
                                    object hostApplcation,
                                    ITestBuilder testBuilder = null,
                                    ITestRunner testRunner = null,
                                    ITestSummaryFormatter testSummaryFormatter = null,
                                    ITestResultCollector externalTestResultCollector = null)
        {
            var applicationHelper = new AccessApplicationHelper(hostApplcation);
            return AccessTestSuite(applicationHelper, testBuilder, testRunner, testSummaryFormatter, externalTestResultCollector);
        }

        public IAccessTestSuite AccessTestSuite(
                                    IAccessApplicationHelper applicationHelper,
                                    ITestBuilder testBuilder = null,
                                    ITestRunner testRunner = null,
                                    ITestSummaryFormatter testSummaryFormatter = null,
                                    ITestResultCollector externalTestResultCollector = null)
        {

            if (testRunner == null)
                testRunner = new TestRunner(applicationHelper.CurrentVBProject);

            if (testBuilder == null)
                testBuilder = new TestBuilder(applicationHelper);

            if (testSummaryFormatter == null)
                testSummaryFormatter = new TestSummaryFormatter(TestSuiteUserSettings.Current.SeparatorMaxLength, TestSuiteUserSettings.Current.SeparatorChar);

            var testSuite = new AccessTestSuite(applicationHelper, testBuilder, testRunner, testSummaryFormatter);

            if (externalTestResultCollector != null)
                testSuite.TestResultCollector = externalTestResultCollector;

            return testSuite;
        }

        public ICodeCoverageTracker CodeCoverageTracker(object vbProject = null)
        {
            return new CodeCoverageTracker((VBProject)vbProject);
        }

        public IErrorTrappingObserver AccessErrorTrappingObserver(object HostApplication)
        {
            return new AccessErrorTrappingObserver(HostApplication);
        }
    }
}
