namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IConstraint
    {
        IMatchResult Matches(object actual);
        IConstraint Child { get; set; }
    }
}
