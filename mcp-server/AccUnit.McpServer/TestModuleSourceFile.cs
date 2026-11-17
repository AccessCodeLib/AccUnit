using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// A test class or test related module in a source file (VBA export format), checked before it is imported.
    /// VBComponents.Import reads the ANSI code page, while source files of agents and msaccess-vcs are usually UTF-8.
    /// </summary>
    public sealed class TestModuleSourceFile
    {
        // Same header rule as TestClassReader.IsTestRelatedCodeModule in the AccUnit core.
        private static readonly Regex TestModuleHeaderRegex = new Regex(@"^\s*'\s*AccUnit:(TestClass|TestRelated)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ProcedureDeclarationRegex = new Regex(
            @"^\s*((Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+\w+",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        private static readonly Regex VbNameRegex = new Regex(@"^\s*Attribute\s+VB_Name\s*=\s*""([^""]+)""",
            RegexOptions.CultureInvariant | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ClassExportHeaderRegex = new Regex(@"\A\s*VERSION\s+1\.0\s+CLASS\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        private const string ClassModuleExtension = ".cls";
        private static readonly string[] LineSeparators = { "\r\n", "\n" };
        private static readonly byte[] Utf8Preamble = Encoding.UTF8.GetPreamble();

        private TestModuleSourceFile(string filePath, string sourceCode)
        {
            FilePath = filePath;
            ModuleName = Path.GetFileNameWithoutExtension(filePath);
            SourceCode = sourceCode;
        }

        public string FilePath { get; }

        /// <summary>Module name from the file name; the core removes the existing module with this name before the import.</summary>
        public string ModuleName { get; }

        public string SourceCode { get; }

        public static TestModuleSourceFile Read(string filePath, Encoding ansiEncoding)
        {
            return new TestModuleSourceFile(filePath, Decode(File.ReadAllBytes(filePath), ansiEncoding));
        }

        /// <summary>UTF-8 with or without BOM if the bytes are valid UTF-8, otherwise the ANSI code page.</summary>
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

        /// <returns>null if the file can be imported, otherwise the reason why not.</returns>
        public string GetRejectReason(Encoding ansiEncoding)
        {
            if (!HasTestModuleHeader(SourceCode))
                return "Not a test module: the header needs the comment 'AccUnit:TestClass or 'AccUnit:TestRelated.";

            var vbName = GetVbName(SourceCode);
            if (vbName != null && !string.Equals(vbName, ModuleName, StringComparison.OrdinalIgnoreCase))
                return "The file name must match the module name (Attribute VB_Name = \"" + vbName + "\").";

            if (!CanBeEncoded(SourceCode, ansiEncoding))
                return "The file contains characters that are not available in the ANSI code page " + ansiEncoding.WebName +
                       " used by the VBA editor.";

            return null;
        }

        /// <summary>Writes the source in the ANSI code page with CRLF line endings, as the VBA editor exports it.</summary>
        public void WriteAnsiCopy(string directory, Encoding ansiEncoding)
        {
            File.WriteAllText(Path.Combine(directory, Path.GetFileName(FilePath)), ToCrLf(GetImportSource()), ansiEncoding);
        }

        /// <summary>
        /// Source as VBComponents.Import expects it. Code without export header (e.g. the draft of accunit_generate_test_class)
        /// gets the header the VBA editor writes; without it a .cls file would not be imported as class module.
        /// </summary>
        public string GetImportSource()
        {
            if (IsClassModule)
                return ClassExportHeaderRegex.IsMatch(SourceCode) ? SourceCode : GetClassExportHeader(ModuleName) + SourceCode;
            return GetVbName(SourceCode) != null ? SourceCode : "Attribute VB_Name = \"" + ModuleName + "\"\r\n" + SourceCode;
        }

        private bool IsClassModule
        {
            get { return string.Equals(Path.GetExtension(FilePath), ClassModuleExtension, StringComparison.OrdinalIgnoreCase); }
        }

        private static string GetClassExportHeader(string moduleName)
        {
            return "VERSION 1.0 CLASS\r\n" +
                   "BEGIN\r\n" +
                   "  MultiUse = -1  'True\r\n" +
                   "END\r\n" +
                   "Attribute VB_Name = \"" + moduleName + "\"\r\n" +
                   "Attribute VB_GlobalNameSpace = False\r\n" +
                   "Attribute VB_Creatable = False\r\n" +
                   "Attribute VB_PredeclaredId = False\r\n" +
                   "Attribute VB_Exposed = False\r\n";
        }

        public static string ToCrLf(string sourceCode)
        {
            return string.Join("\r\n", sourceCode.Split(LineSeparators, StringSplitOptions.None));
        }

        public static bool HasTestModuleHeader(string sourceCode)
        {
            var headerLines = sourceCode.Split(LineSeparators, StringSplitOptions.None)
                                        .TakeWhile(line => !ProcedureDeclarationRegex.IsMatch(line));
            return TestModuleHeaderRegex.IsMatch(string.Join("\n", headerLines));
        }

        public static string GetVbName(string sourceCode)
        {
            var match = VbNameRegex.Match(sourceCode);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static bool CanBeEncoded(string sourceCode, Encoding ansiEncoding)
        {
            return ansiEncoding.GetString(ansiEncoding.GetBytes(sourceCode)) == sourceCode;
        }
    }
}
