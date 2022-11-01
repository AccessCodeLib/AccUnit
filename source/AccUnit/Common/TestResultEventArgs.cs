using System;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Common
{
    public class TestResultEventArgs : EventArgs
    {
        private readonly ITestResult _testResult;

        public TestResultEventArgs(ITestResult testResult)
        {
            _testResult = testResult;
        }

        public ITestResult TestResult
        {
            get { return _testResult; }
        }
    }
}