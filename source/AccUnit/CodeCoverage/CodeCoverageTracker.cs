using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    public class CodeCoverageTracker : IDisposable
    {
        private Dictionary<string, CodeModuleTracker> _codeModules = new Dictionary<string, CodeModuleTracker>();
        private VBProject _vbProject;

        public CodeCoverageTracker(VBProject vbProject)
        {
            _vbProject = vbProject;
        }

        public void Clear(string codeModuleName = null)
        { 
            if (codeModuleName.Length > 0)
            {
                RemoveCodeCoverageTracker(codeModuleName);
            }
            else
            {
                foreach (var key in _codeModules.Keys)
                {
                    RemoveCodeCoverageTracker(key);
                    _codeModules[key].Procedures.Clear();
                }
                _codeModules.Clear();
            }
        }

        public void Add(string codeModuleName)
        {
            if (_codeModules.ContainsKey(codeModuleName))
            {
                _codeModules.Remove(codeModuleName);
            }
            _codeModules.Add(codeModuleName, NewCodeModuleTracker(codeModuleName));
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
            var procedureCode = cmReader.GetProcedureCode(procedure.Name, procedure.ProcKind);
            const string IgnorePattern = @"(?!\s*If\s.*Then\b)(?!\s*Else\b)(?!\s*ElseIf\b)(?!\s*End If\b)" +
                                         @"(?!\s*Select Case\b)(?!\s*Case\b)(?!\s*End Select\b)" +
                                         @"(?!\s*With\b)(?!\s*End With\b)" +
                                         @"(?!\s*For\b)(?!\s*Next\b)" +
                                         @"(?!\s*Do\b)(?!\s*Loop\b)" +
                                         @"(?!\s*While\b)(?!\s*Wend\b)";

            const string pattern = @"^(\d+\s|^\d+$)(?!\s*CodeCoverageTracker\.Track\b)" + IgnorePattern + "(.*)";
            Regex regex = new Regex(pattern, RegexOptions.Singleline);

            string[] procedureLines = procedureCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            for (int lineNumber = 0; lineNumber < procedureLines.Length; lineNumber++)
            {
                var match = regex.Match(procedureLines[lineNumber]);
                if (match.Success)
                {
                    string lineNoCode = match.Groups[1].Value;
                    string codeLine = match.Groups[2].Value;
                    InsertCodeCoverageTrackerLine(codeModule, procedure, lineNoCode, codeLine, lineNumber);
                }
            }
        }

        private void InsertCodeCoverageTrackerLine(CodeModule codeModule, CodeModuleMember procedure, string lineCode, string codeLine, int lineNo)
        {
            int procStartLine = codeModule.ProcBodyLine[procedure.Name, procedure.ProcKind];
            int cmLineNo = procStartLine + lineNo;

            string procName = FormattedProcedureName(procedure);

            codeModule.ReplaceLine(cmLineNo, $"{lineCode.TrimEnd()} CodeCoverageTracker.Track \"{codeModule.Name}\", \"{procName}\", {lineCode.TrimEnd()}:{codeLine}");
        }

        private string FormattedProcedureName(CodeModuleMember procedure)
        {
            string procName = procedure.Name;

            if (procedure.ProcKind == vbext_ProcKind.vbext_pk_Get)
                procName += "_get";
            else if (procedure.ProcKind == vbext_ProcKind.vbext_pk_Let)
                procName += "_let";
            else if (procedure.ProcKind == vbext_ProcKind.vbext_pk_Set)
                procName += "_set";

            return procName;
        }

        private void RemoveCodeCoverageTracker(string codeModuleName)
        {
            var codeModule = _vbProject.VBComponents.Item(codeModuleName).CodeModule;
            var cmReader = new CodeModuleReader(codeModule);
            
            const string pattern = @"^(\d+\s+)CodeCoverageTracker.Track[^:]*:(.*)";
            Regex regex = new Regex(pattern, RegexOptions.Singleline);

            string[] codeLines = cmReader.SourceCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            for (int lineNumber = 0; lineNumber < codeLines.Length; lineNumber++)
            {
                var match = regex.Match(codeLines[lineNumber]);
                if (match.Success)
                {
                    string newLine = match.Groups[1].Value;
                    if (match.Groups.Count > 2 && match.Groups[2].Value.Length>0)
                    {
                        newLine += match.Groups[2].Value;
                    }
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
            var procedureCode = cmReader.GetProcedureCode(procedure.Name, procedure.ProcKind);
            int trackLinesCount = Regex.Matches(procedureCode, @"^\d+\s+CodeCoverageTracker\.Track", RegexOptions.Multiline).Count;
            tracker.Add(FormattedProcedureName(procedure), trackLinesCount);
        }

        public void Track(string codeModulName, string procedureName, int lineNo)
        {
            if (!_codeModules.ContainsKey(codeModulName))
            {
                //Add(codeModulName);
                _codeModules.Add(codeModulName, new CodeModuleTracker(codeModulName));
            }
            _codeModules[codeModulName].Track(procedureName, lineNo);
        }

        public string GetReport()
        {
            const string SeparatorLine = "------------------------------------------";
            var sb = new StringBuilder();
            sb.AppendLine(SeparatorLine);
            sb.AppendLine("Code Coverage Report:");
            sb.AppendLine("---------------------");
            foreach (var key in _codeModules.Keys.OrderBy(k => k))
            {
                sb.AppendLine($"Codemodule {key}:");
                sb.AppendLine(_codeModules[key].GetCoverageProcedureInfo());
                sb.AppendLine(SeparatorLine);
            }
            return sb.ToString();
        }

        #region IDisposable Support

        bool _disposed;

        public void Dispose()
        {
            Dispose(true);
        }

        protected void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    DisposeManagedResources();
                }
                DisposeUnmanagedResources();
            }
            catch
            {
            }
            finally
            {
                
            }

            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            _codeModules = null;
        }

        void DisposeUnmanagedResources()
        {
            _vbProject = null;
        }

        ~CodeCoverageTracker()
        {
            Dispose(false);
        }

        #endregion

    }
}
