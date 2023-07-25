using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.Common.VBIDETools.VbaProjectManagement
{
    public class Module
    {
        private readonly string _name;
        private readonly Func<CodeModule> _getCodeModuleFunc;
        private IEnumerable<string> _commentLinesUntilFirstProcedure;
        private IEnumerable<Member> _methods;
        private StringBuilder _stringBuilder;
        private int _indentLevel;

        public Module(string name)
        {
            _name = name;
        }

        public Module(string name, Func<CodeModule> getCodeModuleFunc)
        {
            _name = name;
            _getCodeModuleFunc = getCodeModuleFunc;
        }

        public string Name
        {
            get { return _name; }
        }

        public IEnumerable<Member> Methods
        {
            get
            {
                if (_methods == null)
                    ReadMethods();

                return _methods;
            }
        }

        private void ReadMethods()
        {
            var codeModule = _getCodeModuleFunc();
            var methods = new List<Member>();

            var startLine = codeModule.CountOfDeclarationLines + 1;
            while (startLine < codeModule.CountOfLines)
            {
                var procedureName = codeModule.get_ProcOfLine(startLine, out vbext_ProcKind procKind);
                if (procKind == vbext_ProcKind.vbext_pk_Proc)
                {
                    var methodDeclaration = GetMethodDeclaration(codeModule, procedureName, procKind);
                    var member = ParseMethod(procedureName, methodDeclaration);
                    member.GetMemberCommentFunc = GetMemberComment;
                    methods.Add(member);
                }
                startLine += codeModule.ProcCountLines[procedureName, procKind];
            }

            _methods = methods;
        }

        private Member ParseMethod(string procedureName, string methodDeclaration)
        {
            var regex = new Regex(@"^\s*(Public |Private |)(Sub|Function|Property Get|Property Set|Property Let) (.*)\((.*)\)$", RegexOptions.IgnoreCase);
            var matches = regex.Matches(methodDeclaration);
            if (matches.Count != 1)
                throw new InvalidOperationException(string.Format("Invalid methodDeclaration \"{0}\".", methodDeclaration));
            var match = matches[0];
            var memberType = GetMemberType(match.Groups[2].Value);
            var isPublic = match.Groups[1].Value == "Public ";
            var parameterList = match.Groups[4].Value;
            return new Member
            {
                Module = this,
                IsPublic = isPublic,
                Type = memberType,
                Name = procedureName,
                ParameterList = parameterList
            };
        }

        private string GetMethodDeclaration(_CodeModule codeModule, string procedureName, vbext_ProcKind procKind)
        {
            var procBodyLineNmb = codeModule.ProcBodyLine[procedureName, procKind];
            var sb = new StringBuilder();
            var line = codeModule.Lines[procBodyLineNmb, 1];
            int cnt = 0;
            while (line.EndsWith("_"))
            {
                cnt++;
                sb.Append(line.Substring(0, line.Length - 2));
                line = codeModule.Lines[procBodyLineNmb + cnt, 1];
            }
            sb.Append(line);
            return sb.ToString();
        }

        private MemberType GetMemberType(string memberToken)
        {
            switch (memberToken)
            {
                case "Sub":
                    return MemberType.Sub;
                case "Function":
                    return MemberType.Function;
                case "Property Get":
                    return MemberType.Getter;
                case "Propery Set":
                    return MemberType.Setter;
                case "Property Let":
                    return MemberType.Letter;
                default:
                    throw new ArgumentOutOfRangeException("memberToken", memberToken,
                                                          string.Format("Unknown member token \"{0}\".", memberToken));
            }
        }

        private string GetMemberComment(string memberName)
        {
            var codeModule = _getCodeModuleFunc();
            var commentsStartLine = codeModule.ProcStartLine[memberName, vbext_ProcKind.vbext_pk_Proc];
            var bodyStartLine = codeModule.ProcBodyLine[memberName, vbext_ProcKind.vbext_pk_Proc];

            return codeModule.Lines[commentsStartLine, bodyStartLine - commentsStartLine];
        }

        public void SetText(string text)
        {
            _indentLevel = 0;
            _stringBuilder = new StringBuilder();
            _stringBuilder.AppendLine(text);
        }

        public void AppendLine()
        {
            EnsureStringBuilderExists();
            _stringBuilder.AppendLine();
        }

        private void EnsureStringBuilderExists()
        {
            if (_stringBuilder == null)
            {
                _stringBuilder = new StringBuilder();
            }
        }

        public void AppendLine(string line)
        {
            EnsureStringBuilderExists();
            _stringBuilder.AppendLine(GetIndent() + line);
        }

        private string GetIndent()
        {
            return new string(' ', 3 * _indentLevel);
        }

        public void Indent()
        {
            _indentLevel++;
        }

        public void Unindent()
        {
            if (_indentLevel == 0)
                throw new InvalidOperationException("Indent level is already 0. Unindent is not possible.");

            _indentLevel--;
        }

        public bool ContainsHeaderTag(string headerTag)
        {
            if (_commentLinesUntilFirstProcedure == null)
                ReadCommentLinesUntilFirstProcedure();
            // ReSharper disable AssignNullToNotNullAttribute
            return _commentLinesUntilFirstProcedure.Any(l => l.Contains(headerTag));
            // ReSharper restore AssignNullToNotNullAttribute
        }

        private void ReadCommentLinesUntilFirstProcedure()
        {
            var declarationLines = new List<string>();
            var codeModule = _getCodeModuleFunc();
            var procName = codeModule.get_ProcOfLine(codeModule.CountOfDeclarationLines + 1, out vbext_ProcKind procKind);
            var sectionEndLine = (procName != null)
                                     ? codeModule.ProcBodyLine[procName, procKind]
                                     : codeModule.CountOfDeclarationLines;

            if (codeModule.CountOfLines > 0) // issue #76
            {
                using (var reader = new StringReader(codeModule.Lines[1, sectionEndLine]))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.TrimStart().StartsWith("'"))
                            declarationLines.Add(line);
                    }
                }
            }
            _commentLinesUntilFirstProcedure = declarationLines;
        }

        public string GetCurrentContent()
        {
            var codeModule = _getCodeModuleFunc();
            return codeModule.Lines[1, codeModule.CountOfLines];
        }

        public string GetNewContent()
        {
            // TODO AccSpec: Streamline GetCurrentContent() and GetNewContent()
            return _stringBuilder.ToString();
        }
    }
}