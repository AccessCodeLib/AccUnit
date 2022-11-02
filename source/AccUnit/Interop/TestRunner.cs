using Microsoft.Vbe.Interop;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("570D48B4-989D-47CD-852F-F6F8AFE6DD14")]
    public interface ITestRunner : Interfaces.ITestRunner
    {
        /*
         * Run(TestClassInstance, "*")                  ... Alle TestMethoden ausführen
         * Run(TestClassInstance, "MethodenName")       ... Nur einen bestimmten Test ausführen
         * TODO: Run(TestClassInstance, "*Filter*Text*") ... Nur Test, die dem Filterausdruck entsprechen, ausführen
         */
        void Run([MarshalAs(UnmanagedType.IDispatch)] object TestFixtureInstance, string TestMethodName = "*", Interfaces.ITestResultCollector TestResultCollector = null);
    }

    [ComVisible(true)]
    [Guid("DBED9DB2-5F34-46A4-87B1-7CB3C4FB94F5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TestRunner")]
    public class TestRunner : AccUnit.TestRunner.VbaTestRunner, ITestRunner
    {
        public TestRunner(VBProject vbProject = null) : base(vbProject)
        {
        }
    }
}
