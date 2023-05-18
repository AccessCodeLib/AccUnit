using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    public class CodeCoverageTracker
    {
        private readonly Dictionary<string, CodeModuleTracker> _codeModules = new Dictionary<string, CodeModuleTracker>();
        private VBProject _vbProject;

        public CodeCoverageTracker(VBProject vbProject)
        {
            _vbProject = vbProject;
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
            var tracker = new CodeModuleTracker(codeModuleName);
            FillCodeModuleProcedures(tracker);
            return tracker;
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
            int trackLinesCount = Regex.Matches(procedureCode, @"^TestSuite\.Track", RegexOptions.Multiline).Count;
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
