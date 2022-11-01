using System;
using AccessCodeLib.AccUnit.Assertions.Interfaces;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class Assertions : IAssertionsBuilder
    {
        public Assertions()
        {
        }

        public IMatchResultCollector MatchResultCollector { get; set; }

        public void That(object actual, IConstraintBuilder constraint, string infoText = null)
        {
            That(actual, (IConstraint)constraint, infoText);
        }

        public void That(object actual, IConstraint constraint, string infoText = null)
        {
            var result = ConvertMatchResult(constraint.Matches(actual));
            result.InfoText = infoText;
            AddResultToMatchResultCollector(result, infoText);
            if (result.Match == false)
            {
                Fail(result);
            }
        }

        protected virtual void Fail(IMatchResult result)
        {
            if (MatchResultCollector != null)
            {
                if (MatchResultCollector.IgnoreFailedMatchAfterAdd)
                    return;
            }
            throw new AssertionException(result.FormattedText, result);
        }

        protected virtual IMatchResult ConvertMatchResult(IMatchResult result)
        {
            return result;
        }

        protected virtual void AddResultToMatchResultCollector(IMatchResult result, string infoText)
        {
            if (MatchResultCollector != null)
            {
                MatchResultCollector.Add(result, infoText);
            }
        }

        #region IDisposable Support

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    DisposeManagedResources();
                }
                DisposeUnmanagedResources();
            }
            catch
            {
            }

            GC.SuppressFinalize(this);
            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            //MatchResultCollector = null;
        }

        void DisposeUnmanagedResources()
        {
            //_hostApplication = null;
        }

        public virtual void Dispose()
        {
            Dispose(true);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        ~Assertions()
        {
            Dispose(false);
        }

        #endregion
    }
}
