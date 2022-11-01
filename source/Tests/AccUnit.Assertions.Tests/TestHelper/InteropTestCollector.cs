using AccessCodeLib.AccUnit.Interop;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    class InteropTestCollector : Interop.IMatchResultCollector
    {
        public IMatchResult Result { get; set; }
        public string InfoText { get; set; }

        public bool IgnoreFailedMatchAfterAdd { get { return true; } }

        public void Add(IMatchResult result, string infoText = null)
        {
            Result = new Interop.MatchResult(result);
            InfoText = infoText;
        }
    }
}
