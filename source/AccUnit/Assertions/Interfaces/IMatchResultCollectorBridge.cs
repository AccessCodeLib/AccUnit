namespace AccessCodeLib.AccUnit.Assertions.Interfaces
{
    interface IMatchResultCollectorBridge : IMatchResultCollector
    {
        IMatchResultCollector MatchResultCollector { get; set; }
    }
}
