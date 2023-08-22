using AccessCodeLib.AccUnit.Assertions.Constraints;
using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class StringConstraintBuilder : ConstraintBuilderBase<string>, IStringConstraintBuilder, IConstraint
    {

        readonly StringComparison _stringComparison = StringComparison.InvariantCulture;

        public StringConstraintBuilder(StringComparison compareMethod = StringComparison.InvariantCulture)  : base(false)
        {
            _stringComparison = compareMethod;
        }

        public new IStringConstraintBuilder EqualTo(string expected)
        {
            AddComparerConstraint("actual = expected", expected, 0);
            return this;
        }

        public new IStringConstraintBuilder LessThan(string expected)
        {
            AddComparerConstraint("actual < expected", expected, -1);
            return this;
        }

        public new IStringConstraintBuilder LessThanOrEqualTo(string expected)
        {
            AddComparerConstraint("actual <= expected", expected, -1, 0);
            return this;
        }

        public new IStringConstraintBuilder GreaterThan(string expected)
        {
            AddComparerConstraint("actual > expected", expected, +1);
            return this;
        }

        public new IStringConstraintBuilder GreaterThanOrEqualTo(string expected)
        {
            AddComparerConstraint("actual >= expected", expected, +1, 0);
            return this;
        }

        public new IStringConstraintBuilder Null
        {
            get
            {
                AddChild(new NullConstraint());
                return this;
            }
        }

        public new IStringConstraintBuilder DBNull
        {
            get
            {
                AddChild(new DBNullConstraint());
                return this;
            }
        }

        public new IStringConstraintBuilder Empty
        {
            get
            {
                AddChild(new EmptyConstraint());
                return this;
            }
        }

        public new IStringConstraintBuilder Not
        {
            get
            {
                AddChild(new NotConstraint());
                return this;
            }
        }

        protected override void AddComparerConstraint(string compareText, object expected, int expectedComparerResult, int expectedComparerResult2)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult, expectedComparerResult2);
                return;
            }

            var newConstraint = new StringComparerConstraint(compareText, (string)expected, expectedComparerResult, _stringComparison);
            AddChild(newConstraint);
        }

        protected override void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult, int expectedComparerResult2)
        {
            var newConstraint =  new ArrayConstraint<string>(compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild(newConstraint);
        }
    }


    public class ConstraintBuilder : ConstraintBuilderBase<object>, IConstraintBuilder, IConstraint
    {
        public ConstraintBuilder()   
        {
        }

        public ConstraintBuilder(bool strict) : base(strict)
        {
        }

        public new IConstraintBuilder EqualTo(object expected)
        {
            AddComparerConstraint("actual = expected", expected, 0);
            return this;
        }

        public new IConstraintBuilder LessThan(object expected)
        {
            AddComparerConstraint("actual < expected", expected, -1);
            return this;
        }

        public new IConstraintBuilder LessThanOrEqualTo(object expected)
        {
            AddComparerConstraint("actual <= expected", expected, -1, 0);
            return this;
        }

        public new IConstraintBuilder GreaterThan(object expected)
        {
            AddComparerConstraint("actual > expected", expected, +1);
            return this;
        }

        public new IConstraintBuilder GreaterThanOrEqualTo(object expected)
        {
            AddComparerConstraint("actual >= expected", expected, +1, 0);
            return this;
        }

        public new IConstraintBuilder Null
        {
            get
            {
                AddChild(new NullConstraint());
                return this;
            }
        }

        public new IConstraintBuilder DBNull
        {
            get
            {
                AddChild(new DBNullConstraint());
                return this;
            }
        }

        public new IConstraintBuilder Empty
        {
            get
            {
                AddChild(new EmptyConstraint());
                return this;
            }
        }

        public new IConstraintBuilder Not
        {
            get
            {
                AddChild(new NotConstraint());
                return this;
            }
        }

    }

    public abstract class ConstraintBuilderBase<T> : IConstraintBuilderBase<T>, IConstraint
    {
        private IConstraint _firstchild;

        private readonly bool _strict = false;

        public ConstraintBuilderBase()
        {
        }

        public ConstraintBuilderBase(bool strict)
        {
            _strict = strict;
        }

        public IConstraintBuilderBase<T> EqualTo(T expected)
        {
            AddComparerConstraint("actual = expected", expected, 0);
            return this;
        }

        public IConstraintBuilderBase<T> LessThan(T expected)
        {
            AddComparerConstraint("actual < expected", expected, -1);
            return this;
        }

        public IConstraintBuilderBase<T> LessThanOrEqualTo(T expected)
        {
            AddComparerConstraint("actual <= expected", expected, -1, 0);
            return this;
        }

        public IConstraintBuilderBase<T> GreaterThan(T expected)
        {
            AddComparerConstraint("actual > expected", expected, +1);
            return this;
        }
        public IConstraintBuilderBase<T> GreaterThanOrEqualTo(T expected)
        {
            AddComparerConstraint("actual >= expected", expected, +1, 0);
            return this;
        }

        protected void AddComparerConstraint(string compareText, T expected, int expectedComparerResult)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult);
                return;
            }

            if ((object)expected is null)
            {
                if (expectedComparerResult == 0)
                {
                    AddChild(new NullConstraint());
                    return;
                }
                else
                {
                    AddChild(new NotConstraint(new NullConstraint()));
                    return;
                }
            }

            Type myType = GetCompareType(expected, _strict);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, _strict);
            AddChild((IConstraint)newConstraint);
        }

        private void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult)
        {
            Type T = expected.GetType().GetElementType();
            Type myType = typeof(ArrayConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult);
            AddChild((IConstraint)newConstraint);
        }

        public static Type Type2Compare(object v, bool strict = false)
        {
            Type T = v.GetType();
            if (!strict)
            {
                if (IsNumeric(T))
                    T = typeof(double);  // should all numeric types be compared as double?
            }

            return T;
        }

        private static Type GetCompareType(object v, bool strict = false)
        {
            Type T = Type2Compare(v, strict);
            return typeof(ComparerConstraint<>).MakeGenericType(T);
        }

        private static bool IsNumeric(Type T)
        {
            return NumericTypes.Contains(Nullable.GetUnderlyingType(T) ?? T);
        }

        private static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(int), typeof(double), typeof(long), typeof(short), typeof(decimal), typeof(float),
            typeof(byte), typeof(uint), typeof(ulong), typeof(ushort), typeof(sbyte)
        };

        protected virtual void AddComparerConstraint(string compareText, object expected, int expectedComparerResult, int expectedComparerResult2)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult, expectedComparerResult2);
                return;
            }

            Type T = expected.GetType();
            Type myType = typeof(ComparerConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, expectedComparerResult2, _strict);
            AddChild((IConstraint)newConstraint);
        }

        protected virtual void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult, int expectedComparerResult2)
        {
            //var newConstraint = new ArrayConstraint(compareText, expected, expectedComparerResult, expectedComparerResult2);

            Type T = expected.GetType().GetElementType();
            Type myType = typeof(ArrayConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild((IConstraint)newConstraint);
        }

        public IConstraintBuilderBase<T> Null
        {
            get
            {
                AddChild(new NullConstraint());
                return this;
            }
        }

        public IConstraintBuilderBase<T> DBNull
        {
            get
            {
                AddChild(new DBNullConstraint());
                return this;
            }
        }

        public IConstraintBuilderBase<T> Empty
        {
            get
            {
                AddChild(new EmptyConstraint());
                return this;
            }
        }

        public IConstraintBuilderBase<T> Not
        {
            get
            {
                AddChild(new NotConstraint());
                return this;
            }
        }

        IConstraint IConstraint.Child { get; set; }

        protected void AddChild(IConstraint constraint)
        {
            if (_firstchild is null)
            {
                _firstchild = constraint;
            }
            else
            {
                _firstchild.Child = constraint;
            }
        }

        IMatchResult IConstraint.Matches(object actual)
        {
            return _firstchild.Matches(actual);
        }
    }
}
