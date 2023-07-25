using System;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class ArrayConstraint<T> : ConstraintBase
    {
        protected Array Expected { get; }
        protected int ExpectedComparerResult;
        protected int ExpectedComparerResult2;
        protected bool UseOr = false;

        public ArrayConstraint(string compareText, Array expected, int expectedComparerResult)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
        }

        public ArrayConstraint(string compareText, Array expected, int expectedComparerResult, int expectedComparerResult2)
        {
            CompareText = compareText;
            Expected = expected;
            ExpectedComparerResult = expectedComparerResult;
            ExpectedComparerResult2 = expectedComparerResult2;
        }

        protected override IMatchResult Compare(object actual)
        {
            if (!(actual is Array))
            {
                return new MatchResult(CompareText, false, "actual is not an array", actual, Expected);
            }

            var actualElementType = actual.GetType().GetElementType();
            if (actualElementType != typeof(T))
            {
                var typeNameOfT = typeof(T).Name;
                return new MatchResult(CompareText, false, "actual is not an array of type " + typeNameOfT, actual, Expected);
            }

            var actualArray = (Array)actual;

            if (actualArray.Rank != Expected.Rank)
            {
                return new MatchResult(CompareText, false, "actual array has rank " + actualArray.Rank + ", expected array has rank " + Expected.Rank, actual, Expected);
            }

            // compare size of actual and expected or each dimension
            for (int i = 0; i < actualArray.Rank; i++)
            {
                if (actualArray.GetLength(i) != Expected.GetLength(i))
                {
                    return new MatchResult(CompareText, false, "actual array has " + actualArray.GetLength(i) + " elements in dimension " + i + ", expected array has " + Expected.GetLength(i) + " elements in dimension " + i, actual, Expected);
                }
            }

            // check each element
            var matchResult = CheckArray(actualArray, Expected);
            if (!matchResult.Match)
            {
                return matchResult;
            }

            return new MatchResult(CompareText, true, "actual array is equal expected array", actual, Expected);
        }

        IMatchResult CheckArray(Array actualArray, Array expectedArray)
        {
            if (actualArray.Rank != expectedArray.Rank)
            {
                return new MatchResult(CompareText, false, "actual array has rank " + actualArray.Rank + ", expected array has rank " + Expected.Rank, actualArray, expectedArray);
            }

            for (int i = 0; i < actualArray.GetLength(0); i++)
            {
                int[] indices = new int[actualArray.Rank];
                indices[0] = i;
                object actualElement = actualArray.GetValue(indices);
                object expectedElement = expectedArray.GetValue(indices);

                if (actualElement.GetType().IsArray && expectedElement.GetType().IsArray)
                {
                    var matchResult = CheckArray((Array)actualElement, (Array)expectedElement);
                    if (!matchResult.Match)
                    {
                        return matchResult;
                    }
                }
                else
                {
                    var comparer = new ComparerConstraint<object>(CompareText, expectedElement, ExpectedComparerResult);
                    var matchResult = comparer.Matches(actualElement);
                    if (!matchResult.Match)
                    {
                        return matchResult;
                    }
                }
            }

            return new MatchResult(CompareText, true, "actual array is equal expected array", actualArray, expectedArray);
        }
    }
}
