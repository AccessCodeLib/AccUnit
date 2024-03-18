using AccessCodeLib.AccUnit.Interfaces;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Integration
{
    [ComVisible(true)]
    [Guid("330B79CF-D77A-47A4-8EF7-E32B75849137")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Interop.Constants.ProgIdLibName + ".TestResult")]
    public class TestResult : ITestResult
    {
        public TestResult(ITestData test)
        {
            Test = test;
        }

        public ITestData Test { get; private set; }

        public bool Executed { get; set; }

        public bool IsError { get; set; }

        public bool IsFailure { get; set; }

        public bool IsIgnored { get; set; }

        public bool IsPassed { get; set; }

        public string Message { get; set; }

        public string Result
        {
            get
            {
                if (IsFailure)
                {
                    return "Failed";
                }
                else if (IsError)
                {
                    return "Error";
                }
                else if (IsIgnored)
                {
                    return "Ignored";
                }
                else if (IsPassed)
                {
                    return "Success";
                }
                else
                {
                    return "Not executed";
                }
            }
        }

        public double ElapsedTime { get; set; }

        public bool Success { get { return IsPassed || IsIgnored; } }
    }
}