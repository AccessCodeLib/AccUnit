using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class ComparerConstraint<T> : ConstraintBase
    {
        protected T Expected { get; }
        protected int ExpectedComparerResult;
        protected int ExpectedComparerResult2;
        protected bool UseOr = false;

        protected readonly bool Strict = false;

        public ComparerConstraint(string compareText, T expected, int expectedComparerResult, bool strict = false)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
            Strict = strict;
        }

        public ComparerConstraint(string compareText, T expected, int expectedComparerResult, int expectedComparerResult2, bool strict = false)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
            ExpectedComparerResult2 = expectedComparerResult2;
            UseOr = true;
            Strict = strict;
        }

        protected override IMatchResult Compare(object actual)
        {
            if (actual is null)
            {
                if (Expected == null)
                {
                    if (ExpectedComparerResult == 0 || (UseOr == true && ExpectedComparerResult2 == 0))
                    {
                        return new MatchResult(CompareText, true, null, actual, Expected);
                    }
                    return new MatchResult(CompareText, false, "actual is Nothing, expected is Nothing", actual, Expected);
                }

                var typeOfValue = Expected?.GetType();
                if (typeOfValue == typeof(string))
                {
                    return new MatchResult(CompareText, false, "actual is vbNullstring and expected is not vbNullString", actual, Expected);
                }

                return new MatchResult(CompareText, false, "actual is Nothing and expected is not Nothing", actual, Expected);
            }

            if (actual == DBNull.Value)
            {
                if (Expected.Equals(DBNull.Value))
                {
                    if (ExpectedComparerResult == 0 || (UseOr == true && ExpectedComparerResult2 == 0))
                    {
                        return new MatchResult(CompareText, true, null, actual, Expected);
                    }
                    return new MatchResult(CompareText, false, "actual is Null, expected is Null", actual, Expected);
                }

                return new MatchResult(CompareText, false, "actual is Null and expected is not Null", actual, Expected);
            }

            // Check type
            var actualType = ConstraintBuilder.Type2Compare(actual, Strict);
            var expectedType = ConstraintBuilder.Type2Compare(Expected, Strict);
            if (actualType != expectedType)
            {
                var returnText = "actual (" + actual.GetType().Name + ") is not of type " + FormattedTypeDescription(expectedType, Strict);
                return new MatchResult(CompareText, false, returnText, actual, Expected);
            }

            // Check value
            var a = (T)Convert.ChangeType(actual, typeof(T));
            var result = Comparer<T>.Default.Compare(a, Expected);
            if (result == ExpectedComparerResult)
            {
                return new MatchResult(CompareText, true, null, actual, Expected);
            }

            if (UseOr && (result == ExpectedComparerResult2))
            {
                return new MatchResult(CompareText, true, null, actual, Expected);
            }

            string compareInfo;
            if (result < 0)
            {
                compareInfo = "actual is less then expected";
            }
            else if (result > 0)
            {
                compareInfo = "actual is greather then expected";
            }
            else
            {
                compareInfo = "actual is equal expected";
            }

            return new MatchResult(CompareText, false, compareInfo, actual, Expected);
        }

        private static string FormattedTypeDescription(Type type, bool strict = false)
        {
            return (!strict && type == typeof(double)) ? "numeric type" : type.Name;
        }
    }
}
