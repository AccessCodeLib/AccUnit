using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.Common.VBIDETools
{
    /// <summary>
    /// Source files contain characters that are not available in the ANSI code page of the VBA editor.
    /// </summary>
    [Serializable]
    public class UnsupportedSourceCharactersException : Exception
    {
        public UnsupportedSourceCharactersException(IEnumerable<string> filePaths, Encoding ansiEncoding)
            : this(filePaths.ToList(), ansiEncoding)
        {
        }

        private UnsupportedSourceCharactersException(IList<string> filePaths, Encoding ansiEncoding)
            : base(CreateMessage(filePaths, ansiEncoding))
        {
            FilePaths = filePaths;
        }

        public IList<string> FilePaths { get; }

        private static string CreateMessage(IEnumerable<string> filePaths, Encoding ansiEncoding)
        {
            return "The following files contain characters that are not available in the ANSI code page " +
                   ansiEncoding.WebName + " used by the VBA editor:\r\n" + string.Join("\r\n", filePaths);
        }
    }
}
