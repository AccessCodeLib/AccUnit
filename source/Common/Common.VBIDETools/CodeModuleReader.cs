using AccessCodeLib.Common.Tools.Logging;
using System.Text;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleReader
    {
        private readonly CodeModule _codeModule;

        public CodeModuleReader(CodeModule codemodule)
        {
            _codeModule = codemodule;
        }

        public string DeclarationLines
        {
            get
            {
                return _codeModule.Lines[1, _codeModule.CountOfDeclarationLines];
            }
        }

        public string GetProcedureHeader(string procedurename, vbext_ProcKind procKind = vbext_ProcKind.vbext_pk_Proc)
        {
            var startLine = _codeModule.ProcStartLine[procedurename, procKind];
            var endLine = _codeModule.ProcBodyLine[procedurename, procKind];
            using (new BlockLogger($"Procedure: {procedurename}, start line: {startLine}, end line: {endLine}"))
            {
                var header = (startLine == endLine ? string.Empty : _codeModule.Lines[startLine, endLine - startLine]);
                //Logger.Log($"header: {header}");
                return header;
            }
        }

        public string SourceCode
        {
            get
            {
                return _codeModule.Lines[1, _codeModule.CountOfLines];
            }
        }

        public CodeModule CodeModule { get { return _codeModule; } }
        public string ComponentName { get { return _codeModule.Parent.Name; } }
        public vbext_ComponentType ComponentType { get { return _codeModule.Parent.Type; } }

        public CodeModuleInfo CodeModuleInfo
        {
            get { return new CodeModuleInfo {Name = _codeModule.Name, Members = Members}; }
        }

        private CodeModuleMemberList _members;
        public CodeModuleMemberList Members
        {
            get
            {
                if (_members == null)
                    FillMembers();
                return _members;
            }
        }

        public void RefreshMemberList()
        {
            FillMembers();
        }

        private void FillMembers()
        {
            _members = new CodeModuleMemberList();
            var currentLine = _codeModule.CountOfDeclarationLines + 1;
            while (currentLine <= _codeModule.CountOfLines)
            {
                // ReSharper disable UseIndexedProperty
                var tempProcName = _codeModule.get_ProcOfLine(currentLine, out vbext_ProcKind tempProcKind);
                // ReSharper restore UseIndexedProperty
                if (tempProcName.Length > 0)
                {
                    var tempProcLine = _codeModule.Lines[_codeModule.ProcBodyLine[tempProcName, tempProcKind], 1].Trim();
                    bool isPublic;
                    if (_codeModule.Parent.Type == vbext_ComponentType.vbext_ct_StdModule)
                    {
                        isPublic = !tempProcLine.Substring(0, 8).Equals("Private ", System.StringComparison.InvariantCultureIgnoreCase);
                    }
                    else
                    {
                        isPublic = tempProcLine.Substring(0, 7).Equals("Public ", System.StringComparison.InvariantCultureIgnoreCase);
                        if (!isPublic)
                            isPublic = tempProcLine.Substring(0, 7).Equals("Friend ", System.StringComparison.InvariantCultureIgnoreCase);
                    }
                    
                    _members.Add(new CodeModuleMember(tempProcName, tempProcKind, isPublic));
                    currentLine = _codeModule.ProcStartLine[tempProcName, tempProcKind] + _codeModule.ProcCountLines[tempProcName, tempProcKind];
                }
                currentLine++;
            }
        }

        public string GetProcedureDeclaration(string procedureName, vbext_ProcKind procKind = vbext_ProcKind.vbext_pk_Proc)
        {
            var procBodyLineNumber = _codeModule.ProcBodyLine[procedureName, procKind];
            var procCountLines = _codeModule.ProcCountLines[procedureName, procKind];
            var currentLineNumber = procBodyLineNumber;
            var lastLineNumber = procBodyLineNumber + procCountLines;
          
            var sb = new StringBuilder();
            string tempLine;
            do
            {
                tempLine = _codeModule.Lines[currentLineNumber, 1].TrimEnd();
                sb.AppendLine(tempLine);
            } while (tempLine.Substring(tempLine.Length-2).Trim().Equals("_") && currentLineNumber++ <= lastLineNumber);

            return sb.ToString().TrimEnd();
        }

        public string GetProcedureCode(string procedureName, vbext_ProcKind procKind = vbext_ProcKind.vbext_pk_Proc)
        {
            var procStartLineNumber = _codeModule.ProcStartLine[procedureName, procKind];
            var procBodyLineNumber = _codeModule.ProcBodyLine[procedureName, procKind];
            var procCountLines = _codeModule.ProcCountLines[procedureName, procKind];
            var procBodyLineCount = procStartLineNumber + procCountLines - procBodyLineNumber;
            return _codeModule.Lines[procBodyLineNumber, procBodyLineCount];
        }

    }
}
