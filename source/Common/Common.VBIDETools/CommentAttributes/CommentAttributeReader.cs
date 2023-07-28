using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class CommentAttributeReader
    {
        private readonly List<Domain> _domains = new List<Domain>();
        private List<CommentAttribute> CommentAttributes { get; set; }

        public List<Domain> Domains
        {
            get { return _domains; }
        }

        public IList<CommentAttribute> Read(string codeLines)
        {
            CommentAttributes = new List<CommentAttribute>();

            var reader = new CommentsBlockReader();
            foreach (var comment in reader.Read(codeLines))
            {
                ParseComment(comment);
            }
            return CommentAttributes;
        }

        private void ParseComment(string comment)
        {
            var lineParser = new CommentAttributeLineParser();
            var parsedLine = lineParser.Parse(comment);

            if (!parsedLine.FoundAttribute)
                return;

            var domain = Domains.Find(d => d.Identifier == parsedLine.DomainIdentifier);
            if (domain is null)
                return;

            var attributeDefinition = domain.AttributeDefinitions.Find(ad => ad.Identifier == parsedLine.AttributeIdentifier);
            if (attributeDefinition is null)
                return;

            var attribute = attributeDefinition.CreateAttribute(parsedLine.Parameters);
            CommentAttributes.Add(attribute);
        }
    }
}