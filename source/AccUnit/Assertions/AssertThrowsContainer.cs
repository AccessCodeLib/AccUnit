using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Assertions
{
    internal static class AssertThrowsStore
    {
        //uint baseErrorNumber = 0x800A0000;
        private const int BaseErrorNumber = -2146828288;

        private static int _ExpectedErrorNumber;
        public static int ExpectedErrorNumber
        {
            get
            {
                return _ExpectedErrorNumber;
            }
            set
            {
                _ExpectedErrorNumber = value;
                ErrorNumber = _ExpectedErrorNumber;
                if (ErrorNumber <= 0)
                {
                    return;
                }
                ErrorNumber += BaseErrorNumber;
            }
        }

        private static int ErrorNumber { get; set; }

        internal static bool CompaireTestRunnerException(Exception ex, TestResult testResult)
        {
            if (ExpectedErrorNumber == 0)
                return false;

            if (ex == null)
            {
                testResult.IsSuccess = false;
                testResult.IsFailure = true;
                testResult.Message = "Expected error number " + ExpectedErrorNumber.ToString() + " was not thrown.";
                Clear();
                return true;
            }

            int errorCode;
            if (ex is COMException comException)
            {
                errorCode = Marshal.GetHRForException(comException);
            }
            else if (ex is System.Reflection.TargetInvocationException)
            {
                errorCode = Marshal.GetHRForException(ex.InnerException);
            }
            else
            {
                try
                {
                    errorCode = Marshal.GetHRForException(ex);
                }
                catch
                {
                    errorCode = 0;
                }
            }

            if (errorCode == ErrorNumber || errorCode == ExpectedErrorNumber)
            {
                testResult.IsSuccess = true;
                testResult.IsFailure = false;
                //testResult.Message = "Expected error number " + ExpectedErrorNumber.ToString() + " was thrown.";
            }
            else
            {
                testResult.IsSuccess = false;
                testResult.IsFailure = true;
                testResult.Message = "Expected error number " + ExpectedErrorNumber.ToString() + " was not thrown. Error was: (" + errorCode.ToString() + ") " + ex.ToString();
            }
            Clear();
            return true;
        }

        public static void Clear()
        {
            ExpectedErrorNumber = 0;
        }
    }
}
