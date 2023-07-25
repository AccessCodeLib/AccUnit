using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Integration
{
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

        public bool IsSuccess { get; set; }

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
                else if (IsSuccess)
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

    }
}