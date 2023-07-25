using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit
{
    public class TestClassMemberInfo : ITestClassMemberInfo
    {
        internal delegate void GetParentEventHandler(TestClassMemberInfo sender, ref TestClassInfo parent);
        internal event GetParentEventHandler GetParent;
        private readonly IList<VbMsgBoxResult> _msgBoxResults = new List<VbMsgBoxResult>();

        public TestClassMemberInfo(string name)
        {
            Name = name;
            _testRowFilter = new List<int>();
        }

        public TestClassMemberInfo(string name, string procHeader)
            : this(name)
        {
            ReadProcHeader(procHeader);
        }

        private void ReadProcHeader(string procHeader)
        {
            SetIgnoreStateFromProcHeader(procHeader);
            SetDoAutoRollbackStateFromProcHeader(procHeader);
            Tags = GetTagsFromProcHeader(procHeader);
            ReadMsgBoxResultsFromProcHeader(procHeader);
            ReadShowAsFromProcHeader(procHeader);
        }

        public string Name { get; private set; }
        public TestClassInfo Parent
        {
            get
            {
                TestClassInfo parent = null;
                GetParent?.Invoke(this, ref parent);
                return parent;
            }
        }

        private IgnoreInfo _ignoreInfo;
        public IgnoreInfo IgnoreInfo { get { return _ignoreInfo; } }

        private readonly List<int> _testRowFilter;
        public IList<int> TestRowFilter { get { return _testRowFilter; } }

        public TagList Tags { get; private set; }

        public bool DoAutoRollback { get; private set; }

        private static readonly Regex TagLineRegex = new Regex(@"^\s*'\s*AccUnit:Tags\(([^']*)\)\s*$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static TagList GetTagsFromProcHeader(string procHeader)
        {
            var tagLines = from Match m in TagLineRegex.Matches(procHeader)
                           select m.Groups[1].Value.Trim();

            var tags = new TagList();
            tags.AddRange(from line in tagLines
                          from tagName in line.Split(',', ';', '|')
                          select new TestItemTag(tagName.Trim('"', ' ')));
            return tags;
        }

        private static readonly Regex IgnoreMemberRegex = new Regex(@"^\s*'\s*AccUnit:Ignore(.*)$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private void SetIgnoreStateFromProcHeader(string procHeader)
        {
            var match = IgnoreMemberRegex.Match(procHeader);
            _ignoreInfo.Ignore = match.Success;
            if (!match.Success) return;
            var comment = match.Groups[1].Value;
            if (!string.IsNullOrEmpty(comment))
            {
                _ignoreInfo.Comment = comment.Trim();
            }
        }

        private static readonly Regex DoAutoRollbackMemberRegex = new Regex(@"^\s*'\s*AccUnit:Rollback\s*$",
                                                                            RegexOptions.CultureInvariant |
                                                                            RegexOptions.Multiline |
                                                                            RegexOptions.IgnoreCase);
        private void SetDoAutoRollbackStateFromProcHeader(string procHeader)
        {
            var match = DoAutoRollbackMemberRegex.Match(procHeader);
            DoAutoRollback = match.Success;
        }

        private static readonly Regex ShowAsRegex = new Regex(@"^\s*'\s*AccUnit:ShowAs\s*\(""(.*)""\)$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private void ReadShowAsFromProcHeader(string procHeader)
        {
            var match = ShowAsRegex.Match(procHeader);
            if (match.Success)
            {
                ShowAs = match.Groups[1].Value;
            }
        }

        private static readonly Regex MsgBoxResultsLineRegex = new Regex(@"^\s*'\s*AccUnit:ClickingMsgBox\(([^\)]*)\).*$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private void ReadMsgBoxResultsFromProcHeader(string procHeader)
        {
            using (new BlockLogger())
            {
                //Logger.Log(procHeader);

                var match = MsgBoxResultsLineRegex.Match(procHeader);
                if (!match.Success) return;

                var line = match.Groups[1].Value;
                //Logger.Log(line);
                var results = line.Split(',', ';', '|');

                foreach (var result in results)
                {
                    _msgBoxResults.Add(GetMsgBoxResult(result.Trim()));
                }
            }
        }

        private static VbMsgBoxResult GetMsgBoxResult(string resultString)
        {
            Logger.Log(resultString);
            switch (resultString.ToLower())
            {
                case "vbok":
                    return VbMsgBoxResult.vbOK;
                case "vbcancel":
                    return VbMsgBoxResult.vbCancel;
                case "vbyes":
                    return VbMsgBoxResult.vbYes;
                case "vbno":
                    return VbMsgBoxResult.vbNo;
                case "vbabort":
                    return VbMsgBoxResult.vbAbort;
                case "vbignore":
                    return VbMsgBoxResult.vbIgnore;
                case "vbretry":
                    return VbMsgBoxResult.vbRetry;
            }
            throw new Exception(string.Format("'{0}' is not a valid VbMsgBoxResult", resultString));
        }

        public bool IsMatch(IEnumerable<TestItemTag> tags)
        {
            return Tags != null && Tags.IsMatch(tags);
        }

        private readonly List<ITestRow> _testRows = new List<ITestRow>();
        public List<ITestRow> TestRows { get { return _testRows; } }

        public IList<VbMsgBoxResult> MsgBoxResults { get { return _msgBoxResults; } }

        public string ShowAs { get; private set; }

        public string DisplayName
        {
            get
            {
                return !string.IsNullOrEmpty(ShowAs)
                           ? ShowAs
                           : Name;
            }
        }
    }

    public struct IgnoreInfo
    {
        public bool Ignore;
        public string Comment;
    }

}
