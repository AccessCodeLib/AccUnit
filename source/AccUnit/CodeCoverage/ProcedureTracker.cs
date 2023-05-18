using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.CodeCoverage
{
    internal class ProcedureTracker
    {
        private readonly Dictionary<int, int> _coverage = new Dictionary<int, int>();

        public ProcedureTracker(string moduleName, string procedureName, int totalLineCount)
        {
            ModuleName = moduleName;
            ProcedureName = procedureName;
            TotalLineCount = totalLineCount;
        }

        public string ModuleName { get; private set; }
        public string ProcedureName { get; private set; }
        public int TotalLineCount { get; private set; }

        public void Track(int lineNo)
        {
            if (_coverage.ContainsKey(lineNo))
            {
                _coverage[lineNo]++;
            }
            else
            {
                _coverage.Add(lineNo, 1);
            }
        }

        public void Clear()
        {
            _coverage.Clear();
        }

        public double GetCoverage()
        {
            var trackedLinesCount = _coverage.Keys.Count;
            return (double)trackedLinesCount / TotalLineCount;
        }

        public string GetCoverageInfo()
        {
            var trackedLinesCount = _coverage.Keys.Count;
            return $"{trackedLinesCount} / {TotalLineCount}";
        }

        public string GetCoverageLineInfo()
        {
            var sb = new StringBuilder();
            var maxLineNoLength = TotalLineCount.ToString().Length;

            foreach (int lineNo in _coverage.Keys.OrderBy(k => k))
            {
                string lineNoFormatted = lineNo.ToString().PadLeft(maxLineNoLength, '0');
                sb.AppendLine($"{lineNoFormatted} : {_coverage[lineNo]}");
            }
            return sb.ToString();
        }
    }
}
