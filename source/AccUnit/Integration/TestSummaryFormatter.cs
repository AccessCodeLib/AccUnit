using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Text;

namespace AccessCodeLib.AccUnit
{

    internal static class TestResultText
    {
        public const string Success = "Success";
        public const string Failure = "Failure";
        public const string Ignored = "Ignored";
        public const string Error = "Error";
        public const string FailedMarker = "***";
        public const string IgoreMarker = "~~~";
    }

    internal class TestSummaryFormatter : ITestSummaryFormatter
    {
        private const int DefaultSeparatorMaxLength = 0;
        private const char DefaultSeparatorChar = '-';
        private const int DefaultFixtureFinishedSeparatorLength = 5;
        private const int DefaultTestCaseResultStartPos = 50;

        public TestSummaryFormatter()
            : this(DefaultSeparatorMaxLength, DefaultSeparatorChar)
        {
        }

        public TestSummaryFormatter(int separatorMaxLength, char separatorChar,
                                    int fixtureFinishedSeparatorLength = DefaultFixtureFinishedSeparatorLength,
                                    int testCaseResultStartPos = DefaultTestCaseResultStartPos)
        {
            SeparatorMaxLength = separatorMaxLength;
            SeparatorChar = separatorChar;
            TestFixtureFinishedSeparatorLength = fixtureFinishedSeparatorLength;
            TestCaseResultStartPos = testCaseResultStartPos;
        }

        public int SeparatorMaxLength { get; set; }
        public char SeparatorChar { get; set; }
        public int TestFixtureFinishedSeparatorLength { get; set; }
        public int TestCaseResultStartPos { get; set; }

        private string SeparatorLine { get { return new string(SeparatorChar, SeparatorMaxLength); } }
        private static string CurrentTimeString { get { return DateTime.Now.ToString("dd.MM.yy HH:mm:ss"); } }

        public string GetTestSummaryText(ITestSummary summary)
        {
            const int captionLength = 9;

            var separatorLine = new string(SeparatorChar, captionLength + summary.Total.ToString().Length);
            var timeString = String.Format("Time   : {0} ms", Math.Round(summary.ElapsedTime, 1));

            var sb = new StringBuilder();
            string maxSeparatorLine;
            if (SeparatorMaxLength > 0)
            {
                maxSeparatorLine = new string(SeparatorChar, SeparatorMaxLength);
                sb.AppendLine(maxSeparatorLine);
            }
            else
            {
                maxSeparatorLine = new string(SeparatorChar, Math.Max(timeString.Length, separatorLine.Length));
            }

            sb.AppendLine(String.Format("Total  : {0}", summary.Total));
            sb.AppendLine(separatorLine);
            sb.AppendLine(String.Format("Passed : {0}", summary.Passed));
            sb.AppendLine(String.Format("Failed : {0}", summary.Failed + summary.Error));
            sb.AppendLine(String.Format("Ignored: {0}", summary.Ignored));

            sb.AppendLine(maxSeparatorLine);
            sb.AppendLine(timeString);
            sb.AppendLine(maxSeparatorLine);
            if ((summary.Failed + summary.Error) > 0)
            {
                sb.AppendLine(String.Format("{0} / {1} failed", summary.Failed + summary.Error, summary.Total));
            }
            else if (summary.Passed == summary.Total)
            {
                sb.AppendLine(String.Format("{0} / {1} passed", summary.Passed, summary.Total));
            }
            else
            {
                sb.AppendLine(String.Format("{0} / {1} ignored", summary.Ignored, summary.Total));
            }

            if (SeparatorMaxLength > 0)
            {
                sb.AppendLine(maxSeparatorLine);
            }

            return sb.ToString();
        }

        public string GetTestCaseFinishedText(ITestResult result)
        {
            var sb = new StringBuilder();

            sb.Append(result.Test.FullName);
            if (result.Test.FullName.Length < TestCaseResultStartPos)
                sb.Append(new string(' ', TestCaseResultStartPos - result.Test.Name.Length));
            else
                sb.Append(" ");

            if (result.IsSuccess)
            {
                sb.Append(TestResultText.Success);
            }
            else if (result.IsFailure)
            {
                sb.AppendFormat("{0} {1}", TestResultText.Failure, TestResultText.FailedMarker);
            }
            else if (result.IsIgnored)
            {
                sb.AppendFormat("{0} {1}", TestResultText.Ignored, TestResultText.IgoreMarker);
            }
            else if (result.IsError)
            {
                sb.AppendFormat("{0} {1}", TestResultText.Error, TestResultText.FailedMarker);
            }

            try
            {
                if (result.Message != null)
                {
                    sb.AppendLine().Append(result.Message);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine().Append(ex.Message);
            }

            return sb.ToString();
        }


        public string GetTestFixtureFinishedText(ITestResult result)
        {
            var sb = new StringBuilder();

            sb.AppendLine(new string(SeparatorChar, TestFixtureFinishedSeparatorLength));
            sb.AppendLine(String.Format("Finished: {0}", CurrentTimeString));

            return sb.ToString();
        }

        public string GetTestFixtureStartedText(ITestFixture fixture)
        {
            var sb = new StringBuilder();
            try
            {
                sb.AppendLine(SeparatorLine);
                sb.AppendLine(fixture.Name);
                sb.AppendLine(new string(SeparatorChar, fixture.Name.Length));
                sb.AppendLine(String.Format("Started: {0}", CurrentTimeString));
            }
            catch (Exception ex)
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        public string GetTestSuiteStartedText(ITestSuite suite)
        {
            return suite.Name;
        }

        public string GetTestSuiteFinishedText(ITestResult result)
        {
            var sb = new StringBuilder();

            sb.AppendLine(new string(SeparatorChar, TestFixtureFinishedSeparatorLength));
            sb.AppendLine(String.Format("Finished: {0}", CurrentTimeString));

            return sb.ToString();
        }

    }
}
