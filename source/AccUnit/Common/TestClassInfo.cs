using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.Common
{
    public class TestClassInfo
    {
        public TestClassInfo(string name) 
            : this(name, null, null)
        { }

        public TestClassInfo(string name, TestClassMemberList members)
            : this(name, null, members)
        { }

        public TestClassInfo(string name, string source)
            : this (name, source, null)
        { }

        public TestClassInfo(string name, string source, TestClassMemberList members)
        {
            _name = name;
            _classtags = GetTagsFromSourceCode(source);
            InitMembers(members);
        }

        public void InitMembers(TestClassMemberList members)
        {
            Members = members;
            if (members == null) return;
            foreach (var m in members)
            {
                m.GetParent += OnMembersGetParent;
            }
            if (_tags != null)
            {
                _tags.AddRange(Members.Tags);
            }
        }

        private void OnMembersGetParent(TestClassMemberInfo sender, ref TestClassInfo parent)
        {
            parent = this;
        }

        private readonly string _name;
        public string Name
        {
            get { return _name; }
        }

        public override string ToString() { return Name; }

        public TestClassMemberList Members { get; private set; }

        private TagList _tags;
        private readonly TagList _classtags;
        public TagList Tags
        {
            get
            {
                if (_tags == null)
                    FillTagList();
                return _tags;
            }
        }

        private void FillTagList()
        {
            _tags = new TagList();

            if (_classtags != null)
            {
                _tags.AddRange(_classtags);
            }

            if (Members != null)
            {
                _tags.AddRange(Members.Tags);
            }
        }

        public bool IsMatch(IEnumerable<TestItemTag> tags)
        {
            if (_tags == null && _classtags == null && Members == null)
            {
                return false;
            }
            return Tags.IsMatch(tags);
        }

        /// @todo use enum for AccUnit attributes
        private static readonly Regex TagLineRegex = new Regex(@"^\s*'\s*AccUnit:TestClass:Tags\(([^']*)\)\s*$", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static TagList GetTagsFromSourceCode(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            var tagLines = from Match m in TagLineRegex.Matches(text)
                           select m.Groups[1].Value.Trim();

            var tags = new TagList();
            tags.AddRange(from line in tagLines
                          from tagName in line.Split(',', ';', '|')
                          select new TestItemTag(tagName.Trim('"', ' ')));
            return tags;
        }

    }
}
