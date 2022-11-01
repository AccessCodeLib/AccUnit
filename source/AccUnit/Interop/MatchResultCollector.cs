using AccessCodeLib.AccUnit.Assertions;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("D98ACABB-E0E1-443C-922F-D30F5806BC46")]
    public interface IMatchResultCollector
    {
        void Add(IMatchResult Result, string InfoText = null);
        bool IgnoreFailedMatchAfterAdd { get; }
    }

    public class MatchResultCollectorBridge : Assertions.IMatchResultCollector, IMatchResultCollector
    {
        private readonly IMatchResultCollector _collector;

        public MatchResultCollectorBridge(IMatchResultCollector collector)
        {
            _collector = collector;
        }

        public bool IgnoreFailedMatchAfterAdd { get { return _collector.IgnoreFailedMatchAfterAdd; } }

        public void Add(IMatchResult result, string infoText = null)
        {
            _collector.Add(result, infoText);
        }
    }
}
