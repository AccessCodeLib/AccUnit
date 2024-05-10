using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public static class VbeCodePaneTools
    {
        public static string GetCodeModuleMemberNameFromCodePane(_CodePane codepane)
        {
            return GetCodeModuleMemberNameFromCodePane(codepane, out _);
        }

        public static string GetCodeModuleMemberNameFromCodePane(_CodePane codepane, out vbext_ProcKind procKind)
        {
            var startLine = GetStartLineFromCodePaneSelection(codepane);
            var codemodule = codepane.CodeModule;
            return codemodule.get_ProcOfLine(startLine, out procKind);
        }

        public static int GetStartLineFromCodePaneSelection(_CodePane codepane)
        {
            codepane.GetSelection(out int startLine, out _, out _, out _);
            return startLine;
        }

        public static void InsertText(_CodePane codepane, string text, int startLine)
        {
            var codemodule = codepane.CodeModule;
            codemodule.InsertLines(startLine, text);
        }

        public static string GetSelectedText(_CodePane codepane)
        {
            codepane.GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);

            var text = codepane.CodeModule.Lines[startLine, endLine - startLine + 1];

            if (startLine != endLine)
            {
                var lastLine = codepane.CodeModule.Lines[endLine, 1];
                if (lastLine.Length > endColumn)
                    text = text.Substring(1, text.Length - (lastLine.Length - endColumn));
            }

            if (startColumn > 1)
                text = text.Substring(startColumn);

            return text;
        }

    }
}