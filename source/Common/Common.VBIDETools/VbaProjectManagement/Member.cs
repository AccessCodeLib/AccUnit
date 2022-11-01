using System;
using System.Collections.Generic;
using AccessCodeLib.Common.VBIDETools.CommentAttributes;

namespace AccessCodeLib.Common.VBIDETools.VbaProjectManagement
{
    public class Member
    {
        public Module Module { get; set; }
        public string Name { get; set; }
        public bool IsPublic { get; set; }
        internal Func<string, string> GetMemberCommentFunc { get; set; }
        internal string ParameterList { get; set; }
        public MemberType Type { get; set; }

        public IList<Parameter> Parameters
        {
            get
            {
                if (_parameters == null)
                {
                    _parameters = new ParameterListParser().Parse(ParameterList);
                }
                return _parameters;
            }
        }

        private IEnumerable<CommentAttribute> _commentAttributes;
        private IList<Parameter> _parameters;

        public IEnumerable<CommentAttribute> GetCommentAttributes(IEnumerable<Domain> domains)
        {
            if (_commentAttributes == null)
            {
                ReadCommentAttributes(domains);
            }

            return _commentAttributes;
        }

        private void ReadCommentAttributes(IEnumerable<Domain> domains)
        {
            var memberComment = GetMemberCommentFunc(Name);
            var attributeReader = new CommentAttributeReader();
            attributeReader.Domains.AddRange(domains);
            _commentAttributes = attributeReader.Read(memberComment);
        }
    }
}