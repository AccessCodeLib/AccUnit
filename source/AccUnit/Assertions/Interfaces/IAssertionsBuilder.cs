using System;

namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IAssertionsBuilder : IDisposable
    {
        IMatchResultCollector MatchResultCollector { get; set; }

        void That(object actual, IConstraintBuilder constraint, string InfoText = null);
        void That(object actual, IConstraint constraint, string InfoText = null);
    }
}
