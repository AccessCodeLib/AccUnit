namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    class NullConstraint : ConstraintBase
    {
        public NullConstraint()
        {
            CompareText = "Is null";
        }

        protected override IMatchResult Compare(object actual)
        {
            if (actual is null)
            {
                return new MatchResult(CompareText, true, null, actual, null);
            }
            else
            {
                var info = actual is string s && s.Length == 0 ? "Is not null, actual is an empty string" + VbNullStringHint : "Is not null";
                return new MatchResult(CompareText, false, info, actual, null);
            }
        }
    }
}
