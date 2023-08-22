using AccessCodeLib.AccUnit.Interop;

namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IConstraintBuilderBase<T> : IConstraint
    {
        IConstraintBuilderBase<T> EqualTo(T expected);
        IConstraintBuilderBase<T> LessThan(T expected);
        IConstraintBuilderBase<T> LessThanOrEqualTo(T expected);
        IConstraintBuilderBase<T> GreaterThan(T expected);
        IConstraintBuilderBase<T> GreaterThanOrEqualTo(T expected);
    }

    public interface IConstraintBuilder : IConstraint, IConstraintBuilderBase<object>
    {
        new IConstraintBuilder EqualTo(object expected);
        new IConstraintBuilder LessThan(object expected);
        new IConstraintBuilder LessThanOrEqualTo(object expected);
        new IConstraintBuilder GreaterThan(object expected);
        new IConstraintBuilder GreaterThanOrEqualTo(object expected);

        IConstraintBuilder Null { get; }
        IConstraintBuilder DBNull { get; }
        // IConstraintBuilder Nothing { get; } -> Null in interop anzeigen 
        IConstraintBuilder Empty { get; }

        // Is.EquivalentTo
        // Is.InRange

        IConstraintBuilder Not { get; }
        // Is.All
    }

    public interface IStringConstraintBuilder : IConstraint, IConstraintBuilderBase<string>, IConstraintBuilder
    {
        new IConstraintBuilder EqualTo(string expected);
        new IConstraintBuilder LessThan(string expected);
        new IConstraintBuilder LessThanOrEqualTo(string expected);
        new IConstraintBuilder GreaterThan(string expected);
        new IConstraintBuilder GreaterThanOrEqualTo(string expected);

        new IStringConstraintBuilder Not { get; }
    }

}
