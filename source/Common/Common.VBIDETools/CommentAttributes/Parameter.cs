namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public class Parameter
    {
        private readonly string _text;

        public Parameter(string text)
        {
            _text = text;
        }

        public string Text
        {
            get { return _text; }
        }

        public short GetVb6Integer()
        {
            short value;
            if (!short.TryParse(Text, out value))
                throw new TypeMismatchException();

            return value;
        }

        public int GetVb6Long()
        {
            int value;
            if (!int.TryParse(Text, out value))
                throw new TypeMismatchException();

            return value;
        }

        public string GetString()
        {
            if (Text == "")
                throw new InvalidLiteralStringException();

            if (!Text.StartsWith("\"") || !Text.EndsWith("\""))
                throw new InvalidLiteralStringException();
            var temp = Text.Substring(1, Text.Length - 2);

            var tempWithoutEscapedQuotes = temp.Replace("\"\"", "");
            if (tempWithoutEscapedQuotes.Contains("\""))
                throw new InvalidLiteralStringException();

            return temp.Replace("\"\"", "\"");
        }

        #region Equality members

        public bool Equals(Parameter other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other._text, _text);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof (Parameter)) return false;
            return Equals((Parameter) obj);
        }

        public override int GetHashCode()
        {
            return (_text != null ? _text.GetHashCode() : 0);
        }

        #endregion
    }
}