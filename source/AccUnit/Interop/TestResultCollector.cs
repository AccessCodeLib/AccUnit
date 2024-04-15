using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("06F02DAF-5EBC-42F1-A2C4-8BEF443F54FA")]
    public interface ITestCollector : Interfaces.ITestResultCollector
    {
        new void Add(ITestResult testResult);
    }


    [ComVisible(true)]
    [Guid("14654A65-4377-44BA-961E-19DF332B18FD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITestResultCollectorComEvents))]
    [ProgId("AccUnit.TestResultCollector")]
    public class TestResultCollector : Integration.TestResultCollector, ITestCollector
                                            , INotifyingTestResultCollector, ITestResultCollectorEvents
                                            , ITestResultSummaryPrinter
    {
        public TestResultCollector(ITestSuite test) : base(test)
        {
        }
        
        protected override void RaiseTestTraceMessage(string message, CodeCoverage.ICodeCoverageTracker CodeCoverageTracker)
        {
            TestTraceMessage?.Invoke(message, CodeCoverageTracker as ICodeCoverageTracker);
            base.RaiseTestTraceMessage(message, CodeCoverageTracker);
        }

        public new event TestTraceMessageEventHandler TestTraceMessage;
        
    }

    public delegate void TestTraceMessageEventHandler(string Message, ICodeCoverageTracker CodeCoverageTracker);
}
