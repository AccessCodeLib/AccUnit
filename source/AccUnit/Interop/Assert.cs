using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("7DF6AA14-DCBB-4D66-91E4-C4FB7D6CCF5C")]
    public interface IAssert : AccUnit.Assertions.IAssertionsBuilder
    {
        new IMatchResultCollector MatchResultCollector { get; set; }

        void That(object Actual, IConstraintBuilder Constraint, string InfoText = null);
        new void Dispose();

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
        void IsNothing([MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        void IsNotNothing([MarshalAs(UnmanagedType.IDispatch)] object Actual, string InfoText = null);
        
    }

    [ComVisible(true)]
    [Guid("0F16F260-A02D-4B8A-9E3D-6E24419D2F0C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".Assert")]
    public class Assert : AccUnit.Assertions.Assertions, IAssert
    { 
        new public IMatchResultCollector MatchResultCollector {
            get 
            {
                return (IMatchResultCollector)base.MatchResultCollector;
            }
            set 
            {
                base.MatchResultCollector = new MatchResultCollectorBridge(value);
            }
        }

        public void That(object actual, IConstraintBuilder constraint, string infoText = null)
        {
            base.That(actual, constraint, infoText);
        }

        protected override Assertions.IMatchResult ConvertMatchResult(Assertions.IMatchResult result)
        {
            return new MatchResult(result);
        }

        protected override void AddResultToMatchResultCollector(Assertions.IMatchResult result, string infoText)
        {
            if (MatchResultCollector != null)
            {
                MatchResultCollector.Add(result, infoText);
            }
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
