using AccessCodeLib.AccUnit.Interop;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Assertions
{
    public interface IAssertionsBuilder : IDisposable
    {
        IMatchResultCollector MatchResultCollector { get; set; }

        void That(object actual, IConstraintBuilder constraint, string InfoText = null);
        void That(object actual, IConstraint constraint, string InfoText = null);
    }

    public interface IAssertionsComparerMethods
    {
        void AreEqual(object Expected, object Actual, string InfoText = null);
        void AreNotEqual(object Expected, object Actual, string InfoText = null);
        //void AreSame( [MarshalAs(UnmanagedType.IDispatch)]object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        //void AreNotSame([MarshalAs(UnmanagedType.IDispatch)] object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        void Greater(object Arg1, object Arg2, string InfoText = null);
        void GreaterOrEqual(object Arg1, object Arg2, string InfoText = null);
        void Less(object Arg1, object Arg2, string InfoText = null);
        void LessOrEqual(object Arg1, object Arg2, string InfoText = null);
        void IsTrue(bool Condition, string InfoText = null);
        void IsFalse(bool Condition, string InfoText = null);
        void IsEmpty(object Actual, string InfoText = null);
        void IsNull(object Actual, string InfoText = null);
        void IsNotNull(object Actual, string InfoText = null);
        void IsNothing(object Actual, string InfoText = null);
        void IsNotNothing(object Actual, string InfoText = null);
        void Throws(int ErrorNumber, string InfoText = null);
    }

}
