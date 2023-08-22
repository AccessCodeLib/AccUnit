using System;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class StringComparerConstraint : ComparerConstraint<string>
    {
        readonly StringComparison _compareMethod;

        public StringComparerConstraint(string compareText, string expected, int expectedComparerResult, StringComparison compareMethod = StringComparison.InvariantCulture)
            : base(compareText, expected, expectedComparerResult, false)
        {
            _compareMethod = compareMethod;
        }

        public StringComparerConstraint(string compareText, string expected, int expectedComparerResult, int expectedComparerResult2, StringComparison compareMethod = StringComparison.InvariantCulture)
             : base(compareText, expected, expectedComparerResult, expectedComparerResult2, false)
        {
            _compareMethod = compareMethod;
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
                    return new MatchResult(CompareText, false, "actual is vbNullstring, expected is vbNullstring", actual, Expected);
                }

                return new MatchResult(CompareText, false, "actual is vbNullstring and expected is not vbNullString", actual, Expected);
            }

            if (actual == DBNull.Value)
            {
                return new MatchResult(CompareText, false, "actual is Null and expected is not Null", actual, Expected);
            }

            // Check type
            var actualType = ConstraintBuilder.Type2Compare(actual, Strict);
            var expectedType = typeof(string);
            if (actualType != expectedType)
            {
                var returnText = "actual (" + actual.GetType().Name + ") is not of type String";
                return new MatchResult(CompareText, false, returnText, actual, Expected);
            }

            // Check value
            var a = (string)Convert.ChangeType(actual, typeof(string));
            var result = string.Compare(a, Expected, _compareMethod);

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
