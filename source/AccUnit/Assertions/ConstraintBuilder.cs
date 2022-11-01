using System;
using System.Collections.Generic;
using AccessCodeLib.AccUnit.Assertions.Constraints;
using AccessCodeLib.AccUnit.Interfaces;

namespace AccessCodeLib.AccUnit.Assertions
{
    public class ConstraintBuilder : IConstraintBuilder, IConstraint
    {
        private IConstraint _firstchild;

        public ConstraintBuilder()
        {
        }

        public IConstraintBuilder EqualTo(object expected)
        {
            //AddChild(new EqualConstraint(expected));
            AddComparerConstraint("actual = expected", expected, 0); 
            return this;
        }

        public IConstraintBuilder LessThan(object expected)
        {
            AddComparerConstraint("actual < expected", expected, -1);
            return this;
        }

        public IConstraintBuilder LessThanOrEqualTo(object expected)
        {
            AddComparerConstraint("actual <= expected", expected, -1, 0);
            return this;
        }

        public IConstraintBuilder GreaterThan(object expected)
        {
            AddComparerConstraint("actual > expected", expected, +1);
            return this;
        }
        public IConstraintBuilder GreaterThanOrEqualTo(object expected)
        {
            AddComparerConstraint("actual >= expected", expected, +1, 0);
            return this;
        }

        private void AddComparerConstraint(string compareText, object expected, int expectedComparerResult)
        {
            if (expected == null)
            {
                if  (expectedComparerResult == 0)
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
            
            Type myType = getCompareType(expected);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult);
            AddChild((IConstraint)newConstraint);
        }

        private Type getCompareType(object v)
        {
            Type T = v.GetType();
            /*
            if (IsIntNumeric(T))
                T = typeof(long);
            else */
            if (IsNumeric(T))
                T = typeof(double);

            return typeof(ComparerContraint<>).MakeGenericType(T);
        }

        /*
        private static bool IsIntNumeric(Type T)
        {
            return IntNumericTypes.Contains(Nullable.GetUnderlyingType(T) ?? T);
        }

        private static readonly HashSet<Type> IntNumericTypes = new HashSet<Type>
        {
            typeof(long), typeof(int), typeof(short), typeof(byte), typeof(sbyte),
            typeof(ulong), typeof(uint), typeof(ushort)
        };
        */
        
        private static bool IsNumeric(Type T)
        {
            return NumericTypes.Contains(Nullable.GetUnderlyingType(T) ?? T);
        }
        
        private static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(int), typeof(double), typeof(long), typeof(short), typeof(decimal), typeof(float),
            typeof(byte), typeof(uint), typeof(ulong), typeof(ushort), typeof(sbyte)
        };

        private void AddComparerConstraint(string compareText, object expected, int expectedComparerResult, int expectedComparerResult2)
        {
            Type T = expected.GetType();
            Type myType = typeof(ComparerContraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild((IConstraint)newConstraint);
        }

        public IConstraintBuilder Null
        {
            get {
                AddChild(new NullConstraint());
                return (this);
            }
        }

        public IConstraintBuilder DBNull
        {
            get
            {
                AddChild(new DBNullConstraint());
                return (this);
            }
        }

        public IConstraintBuilder Empty
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public IConstraintBuilder Not
        {
            get {
                AddChild(new NotConstraint());
                return (this);
            }
        }

        IConstraint IConstraint.Child { get; set; }


        private void AddChild(IConstraint constraint)
        {
            if (_firstchild == null)
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
