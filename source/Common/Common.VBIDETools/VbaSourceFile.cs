using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AccessCodeLib.Common.VBIDETools
{
    /// <summary>
    /// Encoding of VBA source files (.cls, .bas).
    /// VBComponent.Export and VBComponents.Import always use the ANSI code page of the system,
    /// source files in a repository are usually UTF-8.
    /// </summary>
    public static class VbaSourceFile
    {
        private static readonly string[] LineSeparators = { "\r\n", "\n" };
        private static readonly byte[] Utf8Preamble = Encoding.UTF8.GetPreamble();

        /// <summary>Code page used by VBComponent.Export and VBComponents.Import (Encoding.Default on .NET Framework).</summary>
        public static Encoding AnsiEncoding
        {
            get { return Encoding.Default; }
        }

        public static Encoding Utf8BomEncoding
        {
            get { return new UTF8Encoding(true); }
        }

        /// <summary>UTF-8 with BOM, valid UTF-8 without BOM, otherwise the ANSI code page.</summary>
        public static string Decode(byte[] bytes, Encoding ansiEncoding)
        {
            if (bytes.Take(Utf8Preamble.Length).SequenceEqual(Utf8Preamble))
                return Encoding.UTF8.GetString(bytes, Utf8Preamble.Length, bytes.Length - Utf8Preamble.Length);

            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return ansiEncoding.GetString(bytes);
            }
        }

        public static string ReadAllText(string filePath, Encoding ansiEncoding)
        {
            return Decode(File.ReadAllBytes(filePath), ansiEncoding);
        }

        public static bool CanBeEncoded(string sourceCode, Encoding ansiEncoding)
        {
            return ansiEncoding.GetString(ansiEncoding.GetBytes(sourceCode)) == sourceCode;
        }

        public static string ToCrLf(string sourceCode)
        {
            return string.Join("\r\n", sourceCode.Split(LineSeparators, StringSplitOptions.None));
        }

        /// <summary>Rewrites a file exported by the VBA editor (ANSI) in the target encoding with CRLF line endings.</summary>
        public static void ConvertExportedFile(string filePath, Encoding targetEncoding, Encoding ansiEncoding)
        {
            var sourceCode = ansiEncoding.GetString(File.ReadAllBytes(filePath));
            File.WriteAllText(filePath, ToCrLf(sourceCode), targetEncoding);
        }

        /// <summary>
        /// Writes the source file in the ANSI code page with CRLF line endings into the target directory,
        /// as VBComponents.Import expects it.
        /// </summary>
        /// <returns>Path of the ANSI copy (same file name as the source file).</returns>
        /// <exception cref="UnsupportedSourceCharactersException">The file contains characters that are not available in the ANSI code page.</exception>
        public static string WriteAnsiCopy(string sourceFilePath, string targetDirectory, Encoding ansiEncoding)
        {
            var sourceCode = ReadAllText(sourceFilePath, ansiEncoding);
            if (!CanBeEncoded(sourceCode, ansiEncoding))
                throw new UnsupportedSourceCharactersException(new[] { sourceFilePath }, ansiEncoding);

            var targetFilePath = Path.Combine(targetDirectory, Path.GetFileName(sourceFilePath));
            File.WriteAllText(targetFilePath, ToCrLf(sourceCode), ansiEncoding);
            return targetFilePath;
        }

        /// <returns>Files that contain characters which are not available in the ANSI code page.</returns>
        public static IEnumerable<string> GetFilesWithUnsupportedCharacters(IEnumerable<string> filePaths, Encoding ansiEncoding)
        {
            return filePaths.Where(filePath => !CanBeEncoded(ReadAllText(filePath, ansiEncoding), ansiEncoding)).ToList();
        }
    }
}
