using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public static class VbeCodePaneTools
    {
        public static string GetCodeModuleMemberNameFromCodePane(_CodePane codepane)
        {
            vbext_ProcKind procKind;
            return GetCodeModuleMemberNameFromCodePane(codepane, out procKind);
        }

        public static string GetCodeModuleMemberNameFromCodePane(_CodePane codepane, out vbext_ProcKind procKind)
        {
            var startLine = GetStartLineFromCodePaneSelection(codepane);
            var codemodule = codepane.CodeModule;
            // ReSharper disable UseIndexedProperty
            return codemodule.get_ProcOfLine(startLine, out procKind);
            // ReSharper restore UseIndexedProperty
        }

        public static int GetStartLineFromCodePaneSelection(_CodePane codepane)
        {
            int startLine, startColumn, endLine, endColumn;
            codepane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
            return startLine;
        }

        public static void InsertText(_CodePane codepane, string text, int startLine)
        {
            var codemodule = codepane.CodeModule;
            codemodule.InsertLines(startLine, text);
        }

        public static string GetSelectedText(_CodePane codepane)
        {
            int startLine, startColumn, endLine, endColumn;
            codepane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

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