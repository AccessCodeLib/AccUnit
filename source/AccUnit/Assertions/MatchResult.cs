using AccessCodeLib.AccUnit.Interfaces;
using System;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class MatchResult : IMatchResult
    {
        public MatchResult(string compareText, bool match, string text, object actual, object expected, string infoText = null)
        {
            CompareText = compareText;
            Match = match;
            Text = text;
            Actual = actual;
            Expected = expected;
            InfoText = infoText;
        }

        public string FormattedText { get { return formatResultText(Text, Actual, Expected, InfoText); } }

        protected static string formatResultText(string text, object actual, object expected, string infoText = null)
        {
            if (text == null)
            {
                return infoText;
            }
            
            var compareText = "Expected: " + convertToString(expected) + " but was: " + convertToString(actual);
            return $"{text} ({compareText})" + (infoText == null ? "" : $", {infoText}");
        }

        protected static string convertToString(object value)
        {
            if (value == null)
                return "Nothing";

            if (value == DBNull.Value)
                return "Null";

            if (value is string)
                return "\"" + value + "\"";

            return value.ToString();
        }

        public string CompareText { get; private set; }
        public string InfoText { get; set; }

        public bool Match { get; private set; }
        public object Actual { get; private set; }
        public object Expected { get; private set; }

        public string Text { get; private set; }
    }
}
