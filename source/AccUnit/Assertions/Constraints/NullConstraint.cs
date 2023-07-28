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
                return new MatchResult(CompareText, false, "Is not null", actual, null);
            }
        }
    }
}
