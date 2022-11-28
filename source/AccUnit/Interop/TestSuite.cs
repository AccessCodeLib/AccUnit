using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("C80C791F-7C12-4CFA-AD63-DBF428BFA10D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestSuiteComEvents
    {
        void TestTraceMessage(string Message);
        /*
        void TestSuiteStarted(ITestSuite TestSuite);
        void TestFixtureStarted(ITestFixture Fixture);
        void TestStarted(ITest TestCase);
        
        void TestFinished(ITestResult Result);
        void TestFixtureFinished(ITestResult Result);
        */
        
        //void TestSuiteReset(ResetMode Mode, bool Cancel);
        // void Disposed([MarshalAs(UnmanagedType.IDispatch)] object sender);
    }
}
