using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Integration
{
    [ComVisible(true)]
    [Guid("E1BB5665-7C46-4ED3-ACD1-25695AD2EA22")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Interop.Constants.ProgIdLibName + ".TestResultCollection")]
    public class TestResultCollection : List<ITestResult>, ITestResultSummary
    {
        public TestResultCollection(ITestData test)
        {
            Test = test;
        }

        private int ExecutedCount { get; set; }
        private int IsErrorCount { get; set; }
        private int IsFailureCount { get; set; }
        private int IsIgnoredCount { get; set; }
        private int IsSuccessCount { get; set; }

        new public void Add(ITestResult testResult)
        {
            base.Add(testResult);
            ExecutedCount++;
            if (testResult.IsError)
            {
                IsErrorCount++;
                IsError = true;
            }
            else if (testResult.IsFailure)
            {
                IsFailureCount++;
                IsFailure = true;
            }
            else if (testResult.IsIgnored)
            {
                IsIgnoredCount++;
                IsIgnored = true;
            }
            else if (testResult.IsSuccess)
            {
                IsSuccessCount++;
            }

            if (IsSuccessCount == ExecutedCount)
            {
                IsSuccess = true;
            }
            else
            {
                IsSuccess = false;
            }
            Message += "\n" +  testResult.Message;
            Time += testResult.Time;
        }

        public ITestResult Item(int index)
        {
           return base[index];
        }
        
        public ITestData Test { get; private set; }

        public bool Executed { get; set; }

        public bool IsError { get; private set; }

        public bool IsFailure { get; private set; }

        public bool IsIgnored { get; private set; }

        public bool IsSuccess { get; private set; }

        public string Message { get; private set; }

        public string Result
        {
            get
            {
                var resultBuilder = new StringBuilder();
                resultBuilder.Append("Executed: " + ExecutedCount);
                resultBuilder.Append(" Success: " + IsSuccessCount);
                resultBuilder.Append(" Failure: " + IsFailureCount);
                resultBuilder.Append(" Error: " + IsErrorCount);
                resultBuilder.Append(" Ignored: " + IsIgnoredCount);
                return resultBuilder.ToString();
            }
        }
        
        public double Time { get; set; }
        
    }
}