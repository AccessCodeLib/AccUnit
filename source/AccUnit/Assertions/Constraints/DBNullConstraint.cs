using System;

namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class DBNullConstraint : ConstraintBase
    {
        public DBNullConstraint()
        {
            CompareText = "Is DBNull";
        }

        protected override IMatchResult Compare(object actual)
        {
            if (actual == DBNull.Value)
            {
                return new MatchResult(CompareText, true, null, actual, DBNull.Value);
            }
            else
            {
                return new MatchResult(CompareText, false, "Is not DBNull", actual, DBNull.Value);
            }
        }
    }
}
