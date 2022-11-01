using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class Domain
    {
        private readonly string _identifier;
        private readonly List<AttributeDefinitionBase> _attributeDefinitions = new List<AttributeDefinitionBase>();

        public Domain(string identifier)
        {
            _identifier = identifier;
        }

        public string Identifier
        {
            get { return _identifier; }
        }

        public List<AttributeDefinitionBase> AttributeDefinitions
        {
            get { return _attributeDefinitions; }
        }
    }
}