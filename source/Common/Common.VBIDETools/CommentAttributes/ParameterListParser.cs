using System.Collections.Generic;
using System.Text;

namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class ParameterListParser
    {
        public IList<Parameter> Parse(string parameterList)
        {
            if (parameterList.Trim().Length == 0)
                return new List<Parameter>();

            var parameters = new List<Parameter>();
            var inString = false;
            var partBuilder = new StringBuilder();
            foreach (var c in parameterList.ToCharArray())
            {
                if (c == ',')
                {
                    if (inString)
                    {
                        partBuilder.Append(c);
                    }
                    else
                    {
                        parameters.Add(new Parameter(partBuilder.ToString().Trim()));
                        partBuilder = new StringBuilder();
                    }
                }
                else
                {
                    partBuilder.Append(c);

                    if (c == '\"')
                    {
                        inString = !inString;
                    }
                }
            }
            if (inString)
                throw new MalformedParameterListException();
            parameters.Add(new Parameter(partBuilder.ToString().Trim()));
            return parameters;
        }
    }
}