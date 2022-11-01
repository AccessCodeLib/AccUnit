using System;
using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class ParsedCommentAttributeLine
    {
        private readonly IList<Parameter> _parameters;

        public ParsedCommentAttributeLine()
        {
            _foundAttribute = false;
        }

        public ParsedCommentAttributeLine(string domainIdentifier, string attributeIdentifier, IList<Parameter> parameterList)
        {
            _domainIdentifier = domainIdentifier;
            _attributeIdentifier = attributeIdentifier;
            _parameters = parameterList;
            _foundAttribute = true;
        }

        private readonly bool _foundAttribute;
        private readonly string _domainIdentifier;
        private readonly string _attributeIdentifier;

        public bool FoundAttribute
        {
            get { return _foundAttribute; }
        }

        public string DomainIdentifier
        {
            get
            {
                EnsureFoundAttribute();
                return _domainIdentifier;
            }
        }

        public string AttributeIdentifier
        {
            get
            {
                EnsureFoundAttribute();
                return _attributeIdentifier;
            }
        }

        public IList<Parameter> Parameters
        {
            get
            {
                EnsureFoundAttribute();
                return _parameters;
            }
        }

        private void EnsureFoundAttribute()
        {
            if (!FoundAttribute)
                throw new InvalidOperationException("You cannot access this property if FoundAttribute == false.");
        }
    }
}