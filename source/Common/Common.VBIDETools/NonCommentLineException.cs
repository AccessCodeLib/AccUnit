using System;

namespace AccessCodeLib.Common.VBIDETools
{
    public class NonCommentLineException : Exception
    {
        private readonly string _nonCommentLine;

        public NonCommentLineException(string nonCommentLine)
            : base(string.Format("\"{0}\" is not a comment line.", nonCommentLine))
        {
            _nonCommentLine = nonCommentLine;
        }

        public string NonCommentLine
        {
            get { return _nonCommentLine; }
        }
    }
}