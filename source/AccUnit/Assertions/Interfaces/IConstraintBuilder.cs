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

        IConstraintBuilderBase<T> Null { get; }
        IConstraintBuilderBase<T> DBNull { get; }
        // IConstraintBuilder Nothing { get; } -> Null in interop anzeigen 
        IConstraintBuilderBase<T> Empty { get; }

        // Is.EquivalentTo
        // Is.InRange

        IConstraintBuilderBase<T> Not { get; }
        // Is.All
    }

    public interface IConstraintBuilder : IConstraint, IConstraintBuilderBase<object>
    {
        new IConstraintBuilder EqualTo(object expected);
        new IConstraintBuilder LessThan(object expected);
        new IConstraintBuilder LessThanOrEqualTo(object expected);
        new IConstraintBuilder GreaterThan(object expected);
        new IConstraintBuilder GreaterThanOrEqualTo(object expected);

        new IConstraintBuilder Null { get; }
        new IConstraintBuilder DBNull { get; }
        // IConstraintBuilder Nothing { get; } -> Null in interop anzeigen 
        new IConstraintBuilder Empty { get; }

        // Is.EquivalentTo
        // Is.InRange

        new IConstraintBuilder Not { get; }
        // Is.All
    }

    public interface IStringConstraintBuilder : IConstraint, IConstraintBuilderBase<string>
    {
        new IStringConstraintBuilder EqualTo(string expected);
        new IStringConstraintBuilder LessThan(string expected);
        new IStringConstraintBuilder LessThanOrEqualTo(string expected);
        new IStringConstraintBuilder GreaterThan(string expected);
        new IStringConstraintBuilder GreaterThanOrEqualTo(string expected);

        new IStringConstraintBuilder Empty { get; }
        new IStringConstraintBuilder Not { get; }
    }

}
