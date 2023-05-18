using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    internal class CodeModuleTracker
    {
        private static readonly Dictionary<string, ProcedureTracker> _procedures = new Dictionary<string, ProcedureTracker>();

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

        public string GetCoverageProcedureInfo()
        {
            var sb = new StringBuilder();
            var maxLineNoLength = TotalProcedureCount.ToString().Length;

            foreach (var key in _procedures.Keys.OrderBy(k => k))
            {
                sb.AppendLine($"{key} : {_procedures[key].GetCoverage()*100.0:0.0%} ({_procedures[key].GetCoverageInfo()})");
            }
            return sb.ToString();
        }

    }
}
