using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Interop
{
    internal class TestSuite
    {
    }

    [ComVisible(true)]
    [Guid("C80C791F-7C12-4CFA-AD63-DBF428BFA10D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestSuiteComEvents
    {
        void TestTraceMessage(string Message);
        void TestSuiteStarted(ITestSuite TestSuite);
        void TestFixtureStarted(ITestFixture Fixture);
        void Testtarted(ITest TestCase);
        void TestFinished(ITestResult Result);
        void TestFixtureFinished(ITestResult Result);
        void TestSuiteReset(ResetMode Mode, bool Cancel);
        void Disposed([MarshalAs(UnmanagedType.IDispatch)] object sender);
    }
}
