using System;
using System.Collections.Generic;
using AccessCodeLib.AccUnit.Assertions.Constraints;
using AccessCodeLib.AccUnit.Interop;

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

        private static bool IsArray(object objectToCheck)
        {
            return objectToCheck != null && objectToCheck.GetType().IsArray;
        }

        private void AddComparerConstraint(string compareText, object expected, int expectedComparerResult)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult);
                return;
            }

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
            
            Type myType = GetCompareType(expected);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult);
            AddChild((IConstraint)newConstraint);
        }

        private void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult)
        {
            //var newConstraint = new ArrayConstraint(compareText, expected, expectedComparerResult);
            
            Type T = expected.GetType().GetElementType();
            Type myType = typeof(ArrayConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult);
            AddChild((IConstraint)newConstraint);
        }

        public static Type Type2Compare(object v)
        {
            Type T = v.GetType();
            if (IsNumeric(T))
                T = typeof(double);  // should all numeric types be compared as double?

            return T;
        }

        private static Type GetCompareType(object v)
        {
            Type T = Type2Compare(v);
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

        private void AddComparerConstraint(string compareText, object expected, int expectedComparerResult, int expectedComparerResult2)
        {
            if (expected is Array expectedArray)
            {
                AddArrayComparerConstraint(compareText, expectedArray, expectedComparerResult, expectedComparerResult2);
                return;
            }

            Type T = expected.GetType();
            Type myType = typeof(ComparerConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild((IConstraint)newConstraint);
        }

        private void AddArrayComparerConstraint(string compareText, Array expected, int expectedComparerResult, int expectedComparerResult2)
        {
            //var newConstraint = new ArrayConstraint(compareText, expected, expectedComparerResult, expectedComparerResult2);
            
            Type T = expected.GetType().GetElementType();
            Type myType = typeof(ArrayConstraint<>).MakeGenericType(T);
            var newConstraint = Activator.CreateInstance(myType, compareText, expected, expectedComparerResult, expectedComparerResult2);
            AddChild((IConstraint)newConstraint);
        }

        public IConstraintBuilder Null
        {
            get {
                AddChild(new NullConstraint());
                return this;
            }
        }

        public IConstraintBuilder DBNull
        {
            get
            {
                AddChild(new DBNullConstraint());
                return this;
            }
        }

        public IConstraintBuilder Empty
        {
            get
            {
                AddChild(new EmptyConstraint());
                return this;
            }
        }

        public IConstraintBuilder Not
        {
            get {
                AddChild(new NotConstraint());
                return this;
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
