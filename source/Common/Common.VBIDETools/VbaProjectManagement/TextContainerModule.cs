using System;
using System.IO;
using System.Text;

namespace AccessCodeLib.Common.VBIDETools.VbaProjectManagement
{
    public class TextContainerModule
    {
        private readonly string _name;
        private readonly string _rawContent;
        private string _text;

        public TextContainerModule(string name, string rawContent)
        {
            _name = name;
            _rawContent = rawContent;
        }

        public string Text
        {
            get
            {
                if (_text is null)
                    ReadText();

                return _text;
            }
        }

        public string Name
        {
            get { return _name; }
        }

        private void ReadText()
        {
            var sb = new StringBuilder();
            using (var reader = new StringReader(_rawContent))
            {
                string rawLine;
                while ((rawLine = reader.ReadLine()) != null)
                {
                    var line = rawLine.TrimStart();
                    if (line != string.Empty)
                    {
                        if (!line.StartsWith("'"))
                            throw new Exception(string.Format("Invalid line \"{0}\" in TextContainerModule.", rawLine));
                        line = line.Substring(1);
                    }
                    sb.AppendLine(line);
                }
            }
            _text = sb.ToString();
        }
    }
}