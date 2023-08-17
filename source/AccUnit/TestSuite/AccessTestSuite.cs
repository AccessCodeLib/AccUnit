using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public interface IAccessTestSuite : IVBATestSuite
    {
    }

    public class AccessTestSuite : VBATestSuite, IAccessTestSuite
    {
        public enum VbaErrorTrapping : short
        {
            BreakOnAllErrors = 0,
            BreakInClassModule = 1,
            BreakOnUnhandledErrors = 2
        }

        private object _hostApplication;
        public override object HostApplication
        {
            get
            {
                return _hostApplication;
            }
            set
            {
                _hostApplication = value;
                ActiveVBProject = GetCurrentVBProject(_hostApplication);
                base.HostApplication = _hostApplication;
            }
        }

        private AccessApplicationHelper _applicationHelper;
        private AccessApplicationHelper ApplicationHelper
        {
            get { return _applicationHelper ?? (_applicationHelper = new AccessApplicationHelper(_hostApplication)); }
        }

        private VBProject GetCurrentVBProject(object app)
        {
            return app is null ? null : ApplicationHelper.CurrentVBProject;
        }

        protected override void OnTestStarted(TestClassMemberInfo testClassMemberInfo)
        {
            TransactionManager = null;

            EnsureErrorTrappingForTests();

            if (testClassMemberInfo.DoAutoRollback)
            {
                TransactionManager = CreateTransactionManager();
                TransactionManager.BeginTrans();
            }
        }

        private DaoTransactionManager CreateTransactionManager()
        {
            return new DaoTransactionManager(_hostApplication);
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

        private short _errorTrappingBeforeRun;
        public override IVBATestSuite Run(IEnumerable<string> methodFilter = null)
        {
            using (new BlockLogger())
            {
                _errorTrappingBeforeRun = GetAccessErrorTrapping();
                EnsureErrorTrappingForTests();

                base.Run(methodFilter);

                if (_errorTrappingBeforeRun != (short)ErrorTrapping)
                    SetAccessErrorTrapping(_errorTrappingBeforeRun);

                return this;
            }
        }

        private void EnsureErrorTrappingForTests()
        {
            if (GetAccessErrorTrapping() != (short)ErrorTrapping)
                SetAccessErrorTrapping((short)ErrorTrapping);
        }

        private VbaErrorTrapping _errorTrapping = VbaErrorTrapping.BreakOnUnhandledErrors;
        public VbaErrorTrapping ErrorTrapping
        {
            get { return _errorTrapping; }
            set { _errorTrapping = value; }
        }

        private const string ErrorTrappingOptionName = "Error Trapping";
        private short GetAccessErrorTrapping()
        {
            var errorTrapping = (short)ApplicationHelper.GetOption(ErrorTrappingOptionName);
            return errorTrapping;
        }

        private void SetAccessErrorTrapping(short errortrapping)
        {
            ApplicationHelper.SetOption(ErrorTrappingOptionName, errortrapping);
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
            TransactionManager = null;
            _applicationHelper = null;
        }

        void DisposeUnmanagedResources()
        {
            _hostApplication = null;
        }

        ~AccessTestSuite()
        {
            Dispose(false);
        }

        #endregion

    }

}