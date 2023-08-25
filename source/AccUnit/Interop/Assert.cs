using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("7DF6AA14-DCBB-4D66-91E4-C4FB7D6CCF5C")]
    public interface IAssert : AccUnit.Assertions.IAssertionsBuilder, IAssertComparerMethods
    {
        new IMatchResultCollector MatchResultCollector { get; set; }

        void That(object Actual, IConstraint Constraint, string InfoText = null);
        new void Dispose();

        IAssertComparerMethods Strict { get; }

        new void AreEqual(object Expected, object Actual, string InfoText = null);
        new void AreNotEqual(object Expected, object Actual, string InfoText = null);
        //void AreSame( [MarshalAs(UnmanagedType.IDispatch)]object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        //void AreNotSame([MarshalAs(UnmanagedType.IDispatch)] object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        new void Greater(object Arg1, object Arg2, string InfoText = null);
        new void GreaterOrEqual(object Arg1, object Arg2, string InfoText = null);
        new void Less(object Arg1, object Arg2, string InfoText = null);
        new void LessOrEqual(object Arg1, object Arg2, string InfoText = null);
        new void IsTrue(bool Condition, string InfoText = null);
        new void IsFalse(bool Condition, string InfoText = null);
        new void IsEmpty(object Actual, string InfoText = null);
        new void IsNull(object Actual, string InfoText = null);
        new void IsNotNull(object Actual, string InfoText = null);
        new void IsNothing([MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        new void IsNotNothing([MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        new void Throws(int ErrorNumber, string InfoText = null);
    }

    [ComVisible(true)]
    [Guid("A13F3E07-8DE6-4670-844B-B71A946D974C")]
    public interface IAssertComparerMethods : AccUnit.Assertions.IAssertionsComparerMethods
    {
        new void AreEqual(object Expected, object Actual, string InfoText = null);
        new void AreNotEqual(object Expected, object Actual, string InfoText = null);
        //void AreSame( [MarshalAs(UnmanagedType.IDispatch)]object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        //void AreNotSame([MarshalAs(UnmanagedType.IDispatch)] object Expected, [MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        new void Greater(object Arg1, object Arg2, string InfoText = null);
        new void GreaterOrEqual(object Arg1, object Arg2, string InfoText = null);
        new void Less(object Arg1, object Arg2, string InfoText = null);
        new void LessOrEqual(object Arg1, object Arg2, string InfoText = null);
    }


    [ComVisible(true)]
    [Guid("0F16F260-A02D-4B8A-9E3D-6E24419D2F0C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".Assert")]
    public class Assert : AccUnit.Assertions.Assertions, IAssert
    {

        public Assert() : base(false)
        {
        }

        public Assert(bool strict) : base(strict)
        {
        }

        new public IMatchResultCollector MatchResultCollector
        {
            get
            {
                return (IMatchResultCollector)base.MatchResultCollector;
            }
            set
            {
                base.MatchResultCollector = new MatchResultCollectorBridge(value);
            }
        }

        public IAssertComparerMethods Strict { get { return new Assert(true); } }

        public void That(object actual, IConstraint constraint, string infoText = null)
        {
            base.That(actual, constraint, infoText);
        }

        protected override Assertions.IMatchResult ConvertMatchResult(Assertions.IMatchResult result)
        {
            return new MatchResult(result);
        }

        protected override void AddResultToMatchResultCollector(Assertions.IMatchResult result, string infoText)
        {
            MatchResultCollector?.Add(result, infoText);
        }

        #region IDisposable Support

        bool _disposed;


        protected override void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    DisposeManagedResources();
                }
                DisposeUnmanagedResources();
            }
            catch
            {
            }
            finally
            {
                base.Dispose(disposing);
            }

            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            //
        }

        void DisposeUnmanagedResources()
        {
            //_hostApplication = null;
        }

        ~Assert()
        {
            Dispose(false);
        }

        #endregion
    }
}
