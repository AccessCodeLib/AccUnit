using System;

namespace AccessCodeLib.AccUnit.Common
{
    internal static class StringTools
    {
        public static string GetInnerXml(string checkString, string tagName)
        {
            return FindSubString(checkString, "<" + tagName + ">", "</" + tagName + ">");
        }

        private static string FindSubString(string source, string startTag, string endTag)
        {
            var startPos = source.IndexOf(startTag, StringComparison.InvariantCultureIgnoreCase);
            if (startPos < 0) return String.Empty;
            startPos += startTag.Length;
            var endPos = source.IndexOf(endTag, startPos, StringComparison.InvariantCultureIgnoreCase);
            return endPos < 0 ? String.Empty : source.Substring(startPos, endPos - startPos).Trim();
        }
    }
}