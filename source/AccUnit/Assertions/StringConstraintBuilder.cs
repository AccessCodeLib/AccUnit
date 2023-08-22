using AccessCodeLib.AccUnit.Assertions.Constraints;
using System;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class StringConstraintBuilder : ConstraintBuilderBase<string>, IStringConstraintBuilder, IConstraint
    {
        readonly StringComparison _stringComparison = StringComparison.InvariantCulture;

        public StringConstraintBuilder(StringComparison compareMethod = StringComparison.InvariantCulture) : base(false)
        {
            _stringComparison = compareMethod;
        }

        public new IConstraintBuilder EqualTo(string expected)
        {
            AddComparerConstraint("actual = expected", expected, 0);
            return this;
        }

        public new IConstraintBuilder LessThan(string expected)
        {
            AddComparerConstraint("actual < expected", expected, -1);
            return this;
        }

        public new IConstraintBuilder LessThanOrEqualTo(string expected)
        {
            AddComparerConstraint("actual <= expected", expected, -1, 0);
            return this;
        }

        public new IConstraintBuilder GreaterThan(string expected)
        {
            AddComparerConstraint("actual > expected", expected, +1);
            return this;
        }

        public new IConstraintBuilder GreaterThanOrEqualTo(string expected)
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

        public new IStringConstraintBuilder Not
        {
            get
            {
                AddChild(new NotConstraint());
                return this;
            }
        }

        IConstraintBuilder IConstraintBuilder.Not => throw new NotImplementedException();

        protected override void AddComparerConstraint(string compareText, object expected, int expectedComparerResult)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult);
                return;
            }

            if (expected is null)
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

            var newConstraint = new StringComparerConstraint(compareText, (string)expected, expectedComparerResult, _stringComparison);
            AddChild(newConstraint);
        }

        protected override void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult)
        {
            var newConstraint = new ArrayConstraint<string>(compareText, expected, expectedComparerResult);
            AddChild(newConstraint);
        }

        protected override void AddComparerConstraint(string compareText, object expected, int expectedComparerResult, int expectedComparerResult2)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult, expectedComparerResult2);
                return;
            }

            var newConstraint = new StringComparerConstraint(compareText, (string)expected, expectedComparerResult, expectedComparerResult2, _stringComparison);
            AddChild(newConstraint);
        }

        protected override void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult, int expectedComparerResult2)
        {
            var newConstraint = new ArrayConstraint<string>(compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild(newConstraint);
        }

        IConstraintBuilder IConstraintBuilder.EqualTo(object expected)
        {
            return EqualTo(expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.LessThan(object expected)
        {
            return LessThan(expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.LessThanOrEqualTo(object expected)
        {
            return LessThanOrEqualTo(expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.GreaterThan(object expected)
        {
            return GreaterThan(expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.GreaterThanOrEqualTo(object expected)
        {
            return GreaterThanOrEqualTo(expected.ToString());
        }

        IConstraintBuilderBase<object> IConstraintBuilderBase<object>.EqualTo(object expected)
        {
            return EqualTo(expected.ToString());
        }

        IConstraintBuilderBase<object> IConstraintBuilderBase<object>.LessThan(object expected)
        {
            return LessThan(expected.ToString());
        }

        IConstraintBuilderBase<object> IConstraintBuilderBase<object>.LessThanOrEqualTo(object expected)
        {
            return LessThanOrEqualTo(expected.ToString());
        }

        IConstraintBuilderBase<object> IConstraintBuilderBase<object>.GreaterThan(object expected)
        {
            return GreaterThan(expected.ToString());
        }

        IConstraintBuilderBase<object> IConstraintBuilderBase<object>.GreaterThanOrEqualTo(object expected)
        {
            return GreaterThanOrEqualTo(expected.ToString());
        }
    }
}