using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    public class CodeCoverageTracker
    {
        private readonly Dictionary<string, CodeModuleTracker> _codeModules = new Dictionary<string, CodeModuleTracker>();
        private readonly VBProject _vbProject;

        public CodeCoverageTracker(VBProject vbProject)
        {
            _vbProject = vbProject;
        }

        public void Clear(string codeModuleName = null)
        { 
            if (codeModuleName == null)
            {
                foreach (var key in _codeModules.Keys)
                {
                    RemoveCodeCoverageTracker(key);
                }
                _codeModules.Clear();
            }
            else
            {
                RemoveCodeCoverageTracker(codeModuleName);
            }
        }

        public void Add(string codeModuleName)
        {
            if (!_codeModules.ContainsKey(codeModuleName))
            {
                _codeModules.Add(codeModuleName, NewCodeModuleTracker(codeModuleName));
            }
        }

        private CodeModuleTracker NewCodeModuleTracker(string codeModuleName)
        {
            InsertCodeCoverageTracker(codeModuleName);
            var tracker = new CodeModuleTracker(codeModuleName);
            FillCodeModuleProcedures(tracker);
            return tracker;
        }

        private void InsertCodeCoverageTracker(string codeModuleName)
        {
            var codeModule = _vbProject.VBComponents.Item(codeModuleName).CodeModule;
            var cmReader = new CodeModuleReader(codeModule);

            foreach (var procedure in cmReader.Members)
            {
                InsertCodeCoverageTracker(codeModule, cmReader, procedure);
            }
        }

        private void InsertCodeCoverageTracker(CodeModule codeModule, CodeModuleReader cmReader, CodeModuleMember procedure)
        {
            var procedureCode = cmReader.GetProcedureCode(procedure.Name);

            //const string pattern = @"^(\d+:)";
            const string pattern = @"^(\d+:(?!\s*CodeCoverageTracker\.Track\b))";
            Regex regex = new Regex(pattern, RegexOptions.Multiline);

            string[] procedureLines = procedureCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            for (int lineNumber = 0; lineNumber < procedureLines.Length; lineNumber++)
            {
                var match = regex.Match(procedureLines[lineNumber]);
                if (match.Success)
                {
                    string lineCode = match.Groups[1].Value;
                    InsertCodeCoverageTrackerLine(codeModule, procedure, procedureLines[lineNumber], lineCode, lineNumber);
                }
            }
        }

        private void InsertCodeCoverageTrackerLine(CodeModule codeModule, CodeModuleMember procedure, string codeLine, string lineCode, int lineNo)
        {
            int procStartLine = codeModule.ProcBodyLine[procedure.Name, procedure.ProcKind];
            int cmLineNo = procStartLine + lineNo;
            codeModule.ReplaceLine(cmLineNo, $"{lineCode} CodeCoverageTracker.Track \"{codeModule.Name}\", \"{procedure.Name}\", {codeLine}");
        }

        private void RemoveCodeCoverageTracker(string codeModuleName)
        {
            var codeModule = _vbProject.VBComponents.Item(codeModuleName).CodeModule;
            var cmReader = new CodeModuleReader(codeModule);
            
            const string pattern = @"^\d+: CodeCoverageTracker.Track.*(\d+:.*)";
            Regex regex = new Regex(pattern, RegexOptions.Multiline);

            string[] codeLines = cmReader.SourceCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            for (int lineNumber = 0; lineNumber < codeLines.Length; lineNumber++)
            {
                var match = regex.Match(codeLines[lineNumber]);
                if (match.Success)
                {
                    string newLine = match.Groups[1].Value;
                    RemoveCodeCoverageTrackerLine(codeModule, newLine, lineNumber + 1);
                }
            }
        }

        private void RemoveCodeCoverageTrackerLine(CodeModule codeModule, string codeLine, int lineNo)
        {
            codeModule.ReplaceLine(lineNo, codeLine);
        }

        private void FillCodeModuleProcedures(CodeModuleTracker tracker)
        {
            var cm = _vbProject.VBComponents.Item(tracker.CodeModuleName).CodeModule;
            var cmReader = new CodeModuleReader(cm);

            foreach (var procedure in cmReader.Members)
            {
                FillProcedureData(tracker, cmReader, procedure);
            }
        }

        private void FillProcedureData(CodeModuleTracker tracker, CodeModuleReader cmReader, CodeModuleMember procedure)
        {
            var procedureCode = cmReader.GetProcedureCode(procedure.Name);
            int trackLinesCount = Regex.Matches(procedureCode, @"^\d+: CodeCoverageTracker\.Track", RegexOptions.Multiline).Count;
            tracker.Add(procedure.Name, trackLinesCount);
        }

        public void Track(string codeModulName, string procedureName, int lineNo)
        {
            if (!_codeModules.ContainsKey(codeModulName))
            {
                Add(codeModulName);
            }
            _codeModules[codeModulName].Track(procedureName, lineNo);
        }

        public string GetReport()
        {
            const string SeparatorLine = "----";
            var sb = new StringBuilder();
            sb.AppendLine(SeparatorLine);

            sb.AppendLine("Code Coverage Report:");
            foreach (var key in _codeModules.Keys.OrderBy(k => k))
            {
                sb.AppendLine($"Codemodule {key}:");
                sb.AppendLine(_codeModules[key].GetCoverageProcedureInfo());
                sb.AppendLine(SeparatorLine);
            }
            return sb.ToString();
        }

    }
}
