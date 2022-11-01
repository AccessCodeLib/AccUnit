using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class ComparerContraint<T> : ConstraintBase
    {
        protected T Expected { get; }
        protected int ExpectedComparerResult;
        protected int ExpectedComparerResult2;
        protected bool UseOr = false;

        public ComparerContraint(string compareText, T expected, int expectedComparerResult)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
        }

        public ComparerContraint(string compareText, T expected, int expectedComparerResult, int expectedComparerResult2)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
            ExpectedComparerResult2 = expectedComparerResult2;
            UseOr = true;
        }

        protected override IMatchResult Compare(object actual)
        {
            if (actual == null)
            {
                if (Expected == null)
                { 
                    if (ExpectedComparerResult == 0 || UseOr == true && ExpectedComparerResult2 == 0)
                    {
                        return new MatchResult(CompareText, true, null, actual, Expected);
                    }
                    return new MatchResult(CompareText, false, "actual is Nothing, expected is Nothing", actual, Expected);
                }
               
                return new MatchResult(CompareText, false, "actual is Nothing and expected is not Nothing", actual, Expected);
            }
           
            if (actual == DBNull.Value)
            {
                if (Expected.Equals(DBNull.Value))
                {
                    if (ExpectedComparerResult == 0 || UseOr == true && ExpectedComparerResult2 == 0)
                    {
                        return new MatchResult(CompareText, true, null, actual, Expected);
                    }
                    return new MatchResult(CompareText, false, "actual is Null, expected is Null", actual, Expected);
                }

                return new MatchResult(CompareText, false, "actual is Null and expected is not Null", actual, Expected);
            }

            var a = (T)Convert.ChangeType(actual, typeof(T));

            var result = Comparer<T>.Default.Compare(a, Expected);
            //var result = Comparer<IComparable>.Default.Compare((IComparable)actual, (IComparable)Expected);

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
    }
}
