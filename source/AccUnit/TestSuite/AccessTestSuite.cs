using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using System;

namespace AccessCodeLib.AccUnit
{
    public interface IAccessTestSuite : IVBATestSuite
    {
    }

    public class AccessTestSuite : VBATestSuite, IAccessTestSuite
    {
        public AccessTestSuite(IAccessApplicationHelper applicationHelper, IVBATestBuilder testBuilder, ITestRunner testRunner, ITestSummaryFormatter testSummaryFormatter)
                : base(applicationHelper, testBuilder, testRunner, testSummaryFormatter)
        {
        }

        protected new IAccessApplicationHelper ApplicationHelper => (IAccessApplicationHelper)base.ApplicationHelper;

        protected override void OnTestStarted(TestClassMemberInfo testClassMemberInfo)
        {
            TransactionManager = null;

            if (testClassMemberInfo.DoAutoRollback)
            {
                TransactionManager = CreateTransactionManager();
                TransactionManager.BeginTrans();
            }
        }

        private DaoTransactionManager CreateTransactionManager()
        {
            return new DaoTransactionManager(ApplicationHelper.Application);
        }

        private ITransactionManager TransactionManager { get; set; }

        protected override void OnTestFinished(ITestResult result)
        {
            if (TransactionManager is null) return;

            try
            {
                TransactionManager.Rollback();
            }
            catch (Exception xcp)
            {
                Logger.Log(xcp.Message);
                throw;
            }
            finally
            {
                TransactionManager = null;
            }
        }

        public override IVBATestSuite Run()
        {
            using (new BlockLogger())
            {
                using (new AccessErrorTrappingObserver(ApplicationHelper, ErrorTrapping))
                {
                    base.Run();
                }
                return this;
            }
        }

        private VbaErrorTrapping _errorTrapping = VbaErrorTrapping.BreakOnUnhandledErrors;
        public VbaErrorTrapping ErrorTrapping
        {
            get { return _errorTrapping; }
            set { _errorTrapping = value; }
        }

        public bool CheckAccessApplicationIsCompiled()
        {
            return ApplicationHelper.IsCompiled;
        }

        public new IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData)
        {
            _errorTrapping = VbaErrorTrapping.BreakOnUnhandledErrors;
            base.Reset(mode);
            return this;
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
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
            finally
            {
                base.Dispose(disposing);
            }

            GC.SuppressFinalize(this);
            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            // ...
        }

        void DisposeUnmanagedResources()
        {
            TransactionManager = null;
        }

        ~AccessTestSuite()
        {
            Dispose(false);
        }

        #endregion

    }

}