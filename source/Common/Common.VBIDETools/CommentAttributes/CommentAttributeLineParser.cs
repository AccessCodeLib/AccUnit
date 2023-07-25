using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class CommentAttributeLineParser
    {
        private readonly Regex _regex = new Regex(@"^((\w+):|[:@])(\w*)\s*(.*)$", RegexOptions.IgnoreCase);

        public ParsedCommentAttributeLine Parse(string comment)
        {
            var matches = _regex.Matches(comment);
            if (matches.Count == 0)
                return new ParsedCommentAttributeLine();

            var match = matches[0];
            var domainIdentifierGroup = match.Groups[1];
            var namedDomainIdentifierGroup = match.Groups[2];
            var attributeIdentifierGroup = match.Groups[3];
            var parameterListGroup = match.Groups[4];

            string domainIdentifier;
            if (namedDomainIdentifierGroup.Success)
            {
                if (!namedDomainIdentifierGroup.Value.StartsWithLetter())
                    throw new InvalidDomainIdentifierException();
                domainIdentifier = namedDomainIdentifierGroup.Value;
            }
            else
            {
                if (domainIdentifierGroup.Value.StartsWith(":"))
                    throw new MissingDomainIdentifierException();
                domainIdentifier = domainIdentifierGroup.Value;
            }

            var attributeIdentifier = attributeIdentifierGroup.Value;
            if (!attributeIdentifier.StartsWithLetter())
                throw new InvalidIdentifierException();

            var parameters = GetParameterList(parameterListGroup.Value);
            var parsedLine = new ParsedCommentAttributeLine(domainIdentifier, attributeIdentifier, parameters);

            return parsedLine;
        }

        private IList<Parameter> GetParameterList(string parameterList)
        {
            parameterList = parameterList.Trim();

            if (parameterList.StartsWith("(") != parameterList.EndsWith(")"))
                throw new MalformedParameterListException();

            if (parameterList.StartsWith("("))
            {
                parameterList = parameterList.Substring(1, parameterList.Length - 2);
            }

            var parser = new ParameterListParser();

            return parser.Parse(parameterList);
        }
    }
}