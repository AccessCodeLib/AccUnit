using System;
using System.Collections;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    // NUnit-Compare: The actual value must be not-null, a string, Guid, have an int Count property, IEnumerable or DirectoryInfo. The value passed was of type System.Object.
    class EmptyConstraint : ConstraintBase
    {
        public EmptyConstraint()
        {
            CompareText = "Is Empty";
        }

        protected override IMatchResult Compare(object actual)
        {
            // if actual == vbNullString in VBA, object actual == null!
            if (actual is string str)
            {
                if (string.IsNullOrEmpty(str))
                {
                    return new MatchResult(CompareText, true, null, actual, string.Empty);
                }
                else
                {
                    return new MatchResult(CompareText, false, "Is not Empty", actual, string.Empty);
                }
            }
            else if (actual is null)
            {
                return new MatchResult(CompareText, true, null, actual, string.Empty);
            }
            else if (actual == DBNull.Value)
            {
                return new MatchResult(CompareText, false, "Is Null, not Empty", actual, null);
            }
            else if (actual is Array array)
            {
                if (array.Length == 0)
                {
                    return new MatchResult(CompareText, true, null, actual, null);
                }
                else
                {
                    return new MatchResult(CompareText, false, "Is not Empty", actual, null);
                }
            }
            else if (actual is IEnumerable enumerable)
            {
                if (enumerable.GetEnumerator().MoveNext())
                {
                    return new MatchResult(CompareText, false, "Is not Empty", actual, null);
                }
                else
                {
                    return new MatchResult(CompareText, true, null, actual, null);
                }
            }
            else
            {
                return new MatchResult(CompareText, false, "Is not Empty", actual, null);
            }
        }
    }

}
