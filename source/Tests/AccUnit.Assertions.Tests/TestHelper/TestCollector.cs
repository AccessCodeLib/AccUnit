namespace AccessCodeLib.AccUnit.Assertions.Tests
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
