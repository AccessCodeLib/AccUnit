using AccessCodeLib.AccUnit.Assertions;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    class TestCollector : IMatchResultCollector
    {
        public IMatchResult Result { get; set; }
        public string InfoText = string.Empty;

        public bool IgnoreFailedMatchAfterAdd { get { return true; } }

        public void Add(IMatchResult result, string infoText)
        {
            Result = result;
            InfoText = infoText;
        }
    }
}
