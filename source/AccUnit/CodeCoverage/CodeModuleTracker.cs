using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    internal class CodeModuleTracker
    {
        private readonly Dictionary<string, ProcedureTracker> _procedures = new Dictionary<string, ProcedureTracker>();

        public CodeModuleTracker(string codeModuleName)
        {
            CodeModuleName = codeModuleName;
        }

        public string CodeModuleName { get; private set; }

        public void Add(string procedureName, int TotalLineCount)
        {
            if (!_procedures.ContainsKey(procedureName))
            {
                _procedures.Add(procedureName, new ProcedureTracker(CodeModuleName, procedureName, TotalLineCount));
            }
        }

        public void Track(string procedureName, int lineNo)
        {
            if (_procedures.ContainsKey(procedureName))
            {
                _procedures[procedureName].Track(lineNo);
            }
            else
            {
                Add(procedureName, 0);
                _procedures[procedureName].Track(lineNo);
            }
        }

        public Dictionary<string, ProcedureTracker> Procedures { get { return _procedures; } }

        public double GetCoverage()
        {
            return (double)TrackedProcedureCount / _procedures.Keys.Count;
        }

        public string GetCoverageInfo()
        {
            return $"{TrackedProcedureCount} / {TotalProcedureCount}";
        }

        private int TrackedProcedureCount { get { return _procedures.Keys.Where(k => _procedures[k].GetCoverage() > 0).Count(); } }
        private int TotalProcedureCount { get { return _procedures.Keys.Count; } }

        public string GetCoverageProcedureInfo(string procedureName = "*", bool showCoverageDetails = false)
        {
            var sb = new StringBuilder();
            var maxLineNoLength = TotalProcedureCount.ToString().Length;
            
            var procedureKeys = GetFilteredKeys(procedureName).OrderBy(k => k);

            foreach (var key in procedureKeys)
            {
                sb.AppendLine($"{key} : {_procedures[key].GetCoverage():0.0%} ({_procedures[key].GetCoverageInfo()})");
                if (showCoverageDetails)
                {
                    var CoverageLineInfo = _procedures[key].GetCoverageLineInfo();
                    if (!string.IsNullOrEmpty(CoverageLineInfo))
                    {
                        CoverageLineInfo = "  " + CoverageLineInfo.Replace(Environment.NewLine, Environment.NewLine + "  ");
                        sb.AppendLine(CoverageLineInfo);
                    }
                }
            }
            return sb.ToString();
        }

        private IEnumerable<string> GetFilteredKeys(string procedureName = "*")
        {
            if (procedureName == "*" || procedureName == null)
            {
                return _procedures.Keys;
            }
            else
            {
                return _procedures.Keys.Where(k => k == procedureName);
            }
        }
    }
}
