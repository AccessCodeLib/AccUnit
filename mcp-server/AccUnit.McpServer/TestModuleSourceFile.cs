using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// A test class or test related module in a source file (VBA export format, UTF-8 or ANSI), checked before it is imported.
    /// The encoding (VbaSourceFile) and the conversion to ANSI for VBComponents.Import are done by the AccUnit core;
    /// this class adds the checks of the MCP server and the export header for code without header.
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
            return new TestModuleSourceFile(filePath, VbaSourceFile.ReadAllText(filePath, ansiEncoding));
        }

        /// <returns>null if the file can be imported, otherwise the reason why not.</returns>
        public string GetRejectReason(Encoding ansiEncoding)
        {
            if (!HasTestModuleHeader(SourceCode))
                return "Not a test module: the header needs the comment 'AccUnit:TestClass or 'AccUnit:TestRelated.";

            var vbName = GetVbName(SourceCode);
            if (vbName != null && !string.Equals(vbName, ModuleName, StringComparison.OrdinalIgnoreCase))
                return "The file name must match the module name (Attribute VB_Name = \"" + vbName + "\").";

            // Checked per file, so the other files are still imported (the core rejects the whole import).
            if (!VbaSourceFile.CanBeEncoded(SourceCode, ansiEncoding))
                return "The file contains characters that are not available in the ANSI code page " + ansiEncoding.WebName +
                       " used by the VBA editor.";

            return null;
        }

        /// <summary>
        /// Writes the source as it is imported (with export header) in UTF-8 with BOM and CRLF line endings;
        /// the core converts it to the ANSI code page for VBComponents.Import.
        /// </summary>
        /// <returns>Path of the copy (same file name).</returns>
        public string WriteImportCopy(string directory)
        {
            var copyPath = Path.Combine(directory, Path.GetFileName(FilePath));
            File.WriteAllText(copyPath, VbaSourceFile.ToCrLf(GetImportSource()), VbaSourceFile.Utf8BomEncoding);
            return copyPath;
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
    }
}
