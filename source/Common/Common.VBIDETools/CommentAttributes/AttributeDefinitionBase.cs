using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public abstract class AttributeDefinitionBase
    {
        private readonly string _identifier;
        private readonly int? _requiredParameterCount;

        protected AttributeDefinitionBase(string identifier)
        {
            _identifier = identifier;
        }

        protected AttributeDefinitionBase(string identifier, int requiredParameterCount)
        {
            _identifier = identifier;
            _requiredParameterCount = requiredParameterCount;
        }

        public string Identifier
        {
            get { return _identifier; }
        }

        public int? RequiredParameterCount
        {
            get { return _requiredParameterCount; }
        }

        public CommentAttribute CreateAttribute(IList<Parameter> parameters)
        {
            if (parameters == null)
                throw new ArgumentNullException("parameters");

            if (RequiredParameterCount.HasValue)
            {
                if (parameters.Count != RequiredParameterCount.Value)
                    throw new WrongNumberOfParametersException();
            }

            var attribute = CreateSpecificAttribute(parameters);
            Debug.Assert(attribute != null);

            return attribute;
        }

        protected abstract CommentAttribute CreateSpecificAttribute(IList<Parameter> parameters);
    }
}