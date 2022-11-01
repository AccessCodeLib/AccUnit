namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IConstraintBuilder : IConstraint
    {
        IConstraintBuilder EqualTo(object expected);
        IConstraintBuilder LessThan(object expected);
        IConstraintBuilder LessThanOrEqualTo(object expected);
        IConstraintBuilder GreaterThan(object expected);
        IConstraintBuilder GreaterThanOrEqualTo(object expected);

        IConstraintBuilder Null { get; }
        IConstraintBuilder DBNull { get; }
        // IConstraintBuilder Nothing { get; } -> Null in interop anzeigen 
        IConstraintBuilder Empty { get; }

        // Is.EquivalentTo
        // Is.InRange

        IConstraintBuilder Not { get; }
        // Is.All
    }
}
