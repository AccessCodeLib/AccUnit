namespace AccessCodeLib.AccUnit.CodeCoverage
{
    public interface ICodeCoverageTracker
    {
        void Add(string codeModuleName);
        void Clear(string codeModuleName = null);
        void Dispose();
        string GetReport(string codeModuleName = "*", string procedureName = "*", bool showCoverageDetails = false);
        void Track(string codeModulName, string procedureName, int lineNo);
    }
}