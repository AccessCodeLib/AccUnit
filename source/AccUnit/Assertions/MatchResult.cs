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

        public string FormattedText { get { return FormatResultText(Text, Actual, Expected, InfoText); } }

        protected static string FormatResultText(string text, object actual, object expected, string infoText = null)
        {
            if (text is null)
            {
                return infoText;
            }

            var typeOfValue = expected?.GetType() ?? actual?.GetType();
            var compareText = "Expected: " + ConvertToString(expected, typeOfValue) + " but was: " + ConvertToString(actual, typeOfValue);
            return $"{text} ({compareText})" + (string.IsNullOrEmpty(infoText) ? "" : $", {infoText}");
        }

        protected static string ConvertToString(object value, Type typeOfValue = null)
        {
            if (value is null)
                return typeOfValue == typeof(string) ? "vbNullString" : "Nothing";

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
