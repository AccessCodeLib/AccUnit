namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IMatchResultCollector
    {
        void Add(IMatchResult result, string infoText = null);
        bool IgnoreFailedMatchAfterAdd { get; }
    }
}
