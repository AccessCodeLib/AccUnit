﻿using AccessCodeLib.AccUnit.Assertions.Interfaces;
using System;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class Assertions : IAssertionsBuilder, IAssertionsComparerMethods
    {
        public Assertions()
        {
        }

        public Assertions(bool strict)
        {
            _strict = strict;
        }

        private readonly bool _strict = false;

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

        #region Compare-Methoden
        public void AreEqual(object expected, object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).EqualTo(expected), infoText);
        }

        public void AreNotEqual(object expected, object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).Not.EqualTo(expected), infoText);
        }

        public void Greater(object arg1, object arg2, string infoText = null)
        {
            That(arg2, new ConstraintBuilder(_strict).GreaterThan(arg1), infoText);
        }

        public void GreaterOrEqual(object arg1, object arg2, string infoText = null)
        {
            That(arg2, new ConstraintBuilder(_strict).GreaterThanOrEqualTo(arg1), infoText);
        }

        public void Less(object arg1, object arg2, string infoText = null)
        {
            That(arg2, new ConstraintBuilder(_strict).LessThan(arg1), infoText);
        }

        public void LessOrEqual(object arg1, object arg2, string infoText = null)
        {
            That(arg2, new ConstraintBuilder(_strict).LessThanOrEqualTo(arg1), infoText);
        }

        public void IsTrue(bool condition, string infoText = null)
        {
            That(condition, new ConstraintBuilder(_strict).EqualTo(true), infoText);
        }

        public void IsFalse(bool condition, string infoText = null)
        {
            That(condition, new ConstraintBuilder(_strict).EqualTo(false), infoText);
        }

        public void IsEmpty(object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).Empty, infoText);
        }

        public void IsNull(object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).DBNull, infoText);
        }

        public void IsNotNull(object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).Not.DBNull, infoText);
        }

        public void IsNothing(object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).Not.Null, infoText);
        }

        public void IsNotNothing(object actual, string infoText = null)
        {
            That(actual, new ConstraintBuilder(_strict).Not.Null, infoText);
        }

        #endregion

        public void Throws(int ErrorNumber, string InfoText = null)
        {
            AssertThrowsStore.ExpectedErrorNumber = ErrorNumber;
            AssertThrowsStore.InfoText = InfoText;
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
            MatchResultCollector?.Add(result, infoText);
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
