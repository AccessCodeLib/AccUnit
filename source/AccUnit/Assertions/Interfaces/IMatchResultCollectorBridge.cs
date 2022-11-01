using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.Assertions.Interfaces
{
    interface IMatchResultCollectorBridge : IMatchResultCollector
    {
        IMatchResultCollector MatchResultCollector { get; set; }
    }
}
