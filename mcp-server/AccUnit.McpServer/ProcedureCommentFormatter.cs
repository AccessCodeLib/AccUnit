using System;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Extracts the comment above a VBA procedure (e.g. a NaturalDoc block as in the ACLib modules)
    /// from the procedure header lines, i.e. the lines the VBE assigns to the procedure before its declaration.
    /// </summary>
    public static class ProcedureCommentFormatter
    {
        private const string CommentMarker = "'";
        private const string RemMarker = "Rem ";
        private const int MinSeparatorLength = 3;
        private static readonly char[] SeparatorChars = { '-', '=', '*', '#', '_', '~' };

        /// <returns>The comment text without comment markers and separator lines, or null if there is none.</returns>
        public static string Format(string procedureHeader)
        {
            if (string.IsNullOrWhiteSpace(procedureHeader))
                return null;

            var lines = procedureHeader
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                .Select(GetCommentText)
                .Where(text => text != null && !IsSeparator(text))
                .ToList();

            lines = RemoveCommonIndent(lines);
            lines = CollapseEmptyLines(lines);
            return lines.Count == 0 ? null : string.Join(Environment.NewLine, lines);
        }

        /// <returns>The text after the comment marker, or null for code and blank lines.</returns>
        private static string GetCommentText(string line)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith(CommentMarker, StringComparison.Ordinal))
                return trimmed.Substring(CommentMarker.Length).TrimEnd();
            if (trimmed.StartsWith(RemMarker, StringComparison.OrdinalIgnoreCase))
                return trimmed.Substring(RemMarker.Length).TrimEnd();
            return null;
        }

        private static bool IsSeparator(string text)
        {
            var trimmed = text.Trim();
            return trimmed.Length >= MinSeparatorLength && trimmed.All(c => SeparatorChars.Contains(c));
        }

        private static List<string> RemoveCommonIndent(List<string> lines)
        {
            var indents = lines.Where(line => line.Length > 0).Select(line => line.Length - line.TrimStart().Length).ToList();
            var commonIndent = indents.Count == 0 ? 0 : indents.Min();
            return lines.Select(line => line.Length == 0 ? line : line.Substring(commonIndent)).ToList();
        }

        private static List<string> CollapseEmptyLines(List<string> lines)
        {
            var result = new List<string>();
            foreach (var line in lines)
            {
                var previousIsEmpty = result.Count == 0 || result[result.Count - 1].Length == 0;
                if (line.Length == 0 && previousIsEmpty)
                    continue;
                result.Add(line);
            }
            if (result.Count > 0 && result[result.Count - 1].Length == 0)
                result.RemoveAt(result.Count - 1);
            return result;
        }
    }
}
