using System;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class NotConstraint : ConstraintBase
    {
        public NotConstraint()
        {
            CompareText = "Not";
        }

        public NotConstraint(IConstraint child)
        {
            CompareText = "Not";
            Child = child;
        }

        protected override IMatchResult Compare(object actual)
        {
            IMatchResult result;
            if (actual is IMatchResult matchResult)
            {
                result = matchResult;
            }
            else
            {
                throw new NotImplementedException();
            }

            string expectedText = result.Expected is null ? "Nothing" : result.Expected.ToString();
            expectedText = "Not " + expectedText;

            result = !result.Match ? new MatchResult(CompareText, true, null, result.Actual, expectedText)
                                   : new MatchResult(CompareText, false, result.CompareText, result.Actual, expectedText);

            return result;
        }
    }
}
