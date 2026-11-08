using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.TestRunner
{
    /// <summary>
    /// Matches test names against a list of filters.
    /// Supported wildcards (VBA Like style): * (any characters), ? (single character),
    /// [abc] / [a-c] (character class), [!abc] (negated character class).
    /// All other characters are literals. Matching is case-insensitive.
    /// </summary>
    [ComVisible(false)]
    public sealed class TestNameFilterMatcher
    {
        private readonly Regex _regex;

        public TestNameFilterMatcher(IEnumerable<string> filters)
        {
            var alternatives = (filters ?? Enumerable.Empty<string>()).Select(ConvertFilterToRegexPattern);
            var pattern = "^(?:" + string.Join("|", alternatives) + ")$";
            _regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        public static bool ContainsWildcard(string filter)
        {
            return filter != null && filter.IndexOfAny(new[] { '*', '?', '[' }) >= 0;
        }

        public bool IsMatch(string testName)
        {
            return _regex.IsMatch(testName ?? string.Empty);
        }

        private static string ConvertFilterToRegexPattern(string filter)
        {
            var sb = new StringBuilder();
            var text = filter ?? string.Empty;

            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                switch (c)
                {
                    case '*':
                        sb.Append(".*");
                        break;
                    case '?':
                        sb.Append('.');
                        break;
                    case '[':
                        var closePos = FindClassEnd(text, i);
                        if (closePos < 0)
                        {
                            sb.Append(@"\[");
                        }
                        else
                        {
                            sb.Append(ConvertCharacterClass(text.Substring(i + 1, closePos - i - 1)));
                            i = closePos;
                        }
                        break;
                    default:
                        sb.Append(Regex.Escape(c.ToString()));
                        break;
                }
            }

            return sb.ToString();
        }

        // Returns the index of the closing ']' or -1 if the class is unclosed or empty (e.g. "[]" or "[!]").
        private static int FindClassEnd(string text, int openPos)
        {
            var contentStart = openPos + 1;
            if (contentStart < text.Length && text[contentStart] == '!')
            {
                contentStart++;
            }

            var closePos = text.IndexOf(']', contentStart);
            return closePos > contentStart ? closePos : -1;
        }

        private static string ConvertCharacterClass(string content)
        {
            var negate = content.StartsWith("!");
            if (negate)
            {
                content = content.Substring(1);
            }

            // '^' stays as passed through (negation if first), but '\' and '[' must not act as regex syntax
            content = content.Replace(@"\", @"\\").Replace("[", @"\[");
            return "[" + (negate ? "^" : string.Empty) + content + "]";
        }
    }
}
