﻿using AccessCodeLib.AccUnit.Interfaces;
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
    public class TestResultCollection : List<ITestResult>, ITestResultSummary, ITestSummary, ITestResultCollector
    {
        public TestResultCollection(ITestData test)
        {
            Test = test;
        }

        private int ExecutedCount { get; set; }
        private int IsErrorCount { get; set; }
        private int IsFailureCount { get; set; }
        private int IsIgnoredCount { get; set; }
        private int IsPassedCount { get; set; }

        new public void Add(ITestResult testResult)
        {
            base.Add(testResult);

            if (testResult is ITestSummary testSummary)
            {
                ExecutedCount += testSummary.Total;
                if (testResult.IsError)
                {
                    IsErrorCount += testSummary.Error;
                    IsError = true;
                }
                if (testResult.IsFailure)
                {
                    IsFailureCount += testSummary.Failed;
                    IsFailure = true;
                }
                if (testResult.IsIgnored)
                {
                    IsIgnoredCount += testSummary.Ignored;
                    IsIgnored = true;
                }
                IsPassedCount += testSummary.Passed;
            }
            else
            {
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
                else if (testResult.IsPassed)
                {
                    IsPassedCount++;
                }
            }

            if (IsPassedCount == ExecutedCount)
            {
                IsPassed = true;
            }
            else
            {
                IsPassed = false;
            }

            Message += "\n" + testResult.Message;
            ElapsedTime += testResult.ElapsedTime;
        }

        public ITestResult Item(int index)
        {
            return base[index];
        }

        public void Reset()
        {
            Clear();

            ExecutedCount = 0;
            IsErrorCount = 0;
            IsFailureCount = 0;
            IsIgnoredCount = 0;
            IsPassedCount = 0;

            Executed = false;
            IsError = false;
            IsFailure = false;
            IsIgnored = false;
            IsPassed = false;
            Message = string.Empty;

            ElapsedTime = 0;
        }

        public ITestData Test { get; private set; }

        public bool Executed { get; set; }

        public bool IsError { get; private set; }

        public bool IsFailure { get; private set; }

        public bool IsIgnored { get; private set; }

        public bool IsPassed { get; private set; }

        public string Message { get; private set; }

        public string Result
        {
            get
            {
                var resultBuilder = new StringBuilder();
                resultBuilder.Append("Executed: " + ExecutedCount);
                resultBuilder.Append(" Success: " + IsPassedCount);
                resultBuilder.Append(" Failure: " + IsFailureCount);
                resultBuilder.Append(" Error:   " + IsErrorCount);
                resultBuilder.Append(" Ignored: " + IsIgnoredCount);
                return resultBuilder.ToString();
            }
        }

        public double ElapsedTime { get; set; }

        public IEnumerable<ITestResult> TestResults { get { return this; } }
        public ITestResult[] GetTestResults() { return this.ToArray(); }

        public int Total { get { return ExecutedCount; } }

        public int Passed { get { return IsPassedCount; } }

        public int Failed { get { return IsFailureCount; } }

        public int Error { get { return IsErrorCount; } }

        public int Ignored { get { return IsIgnoredCount; } }

        public bool Success { get { return (IsFailureCount + IsErrorCount) == 0; } }
    }
}