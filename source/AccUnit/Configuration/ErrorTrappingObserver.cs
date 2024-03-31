using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Configuration
{
    [ComVisible(true)]
    [Guid("26224F83-8C6A-424B-8F68-830E76917415")]
    public interface IErrorTrappingObserver : IDisposable
    {
        [ComVisible(true)]
        void SetErrorTrapping(VbaErrorTrapping ErrorTrapping);
        new void Dispose(); 
    }

    [ComVisible(true)]
    [Guid("27C88B28-6EDB-4CC2-949A-EEAE25B01A82")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.AccessErrorTrappingObserver")]
    public class AccessErrorTrappingObserver : IErrorTrappingObserver, IDisposable
    {
        private AccessApplicationHelper _applicationHelper;
        private object _hostApplication;
        private readonly VbaErrorTrapping _initialErrorTrapping;

        public AccessErrorTrappingObserver(AccessApplicationHelper applicationHelper)
        {
            _applicationHelper = applicationHelper;
            _initialErrorTrapping = GetAccessErrorTrapping();
        }

        public AccessErrorTrappingObserver(AccessApplicationHelper applicationHelper, VbaErrorTrapping errorTrappingToUse)
        {
            _applicationHelper = applicationHelper;
            _initialErrorTrapping = GetAccessErrorTrapping();
            SetErrorTrapping(errorTrappingToUse);
        }

        public AccessErrorTrappingObserver(object hostApplication)
        {
            _hostApplication = hostApplication;
            _initialErrorTrapping = GetAccessErrorTrapping();
        }

        public AccessErrorTrappingObserver(object hostApplication, VbaErrorTrapping errorTrappingToUse)
        {
            _hostApplication = hostApplication;
            _initialErrorTrapping = GetAccessErrorTrapping();
            SetErrorTrapping(errorTrappingToUse);
        }

        private AccessApplicationHelper ApplicationHelper
        {
            get { return _applicationHelper ?? (_applicationHelper = new AccessApplicationHelper(_hostApplication)); }
        }

        public void SetErrorTrapping(VbaErrorTrapping errorTrapping)
        {
            ApplicationHelper.SetOption(ErrorTrappingOptionName, errorTrapping);
        }

        private const string ErrorTrappingOptionName = "Error Trapping";
        private VbaErrorTrapping GetAccessErrorTrapping()
        {
            var errorTrapping = (VbaErrorTrapping)(short)ApplicationHelper.GetOption(ErrorTrappingOptionName);
            return errorTrapping;
        }

        #region IDisposable Support

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            try
            {
                if (_applicationHelper != null)
                {
                    // set the error trapping back to the initial value
                    if (GetAccessErrorTrapping() != _initialErrorTrapping)
                    {
                        SetErrorTrapping(_initialErrorTrapping);
                    }
                }

                if (disposing)
                {
                    DisposeManagedResources();
                }
                DisposeUnmanagedResources();
                _disposed = true;
            }
            catch (Exception ex) { Logger.Log(ex); }

            GC.Collect();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void DisposeManagedResources()
        {
            _applicationHelper = null;
        }

        void DisposeUnmanagedResources()
        {
            _hostApplication = null;
        }

        ~AccessErrorTrappingObserver()
        {
            Dispose(false);
        }

        #endregion
    }
}
