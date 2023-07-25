using AccessCodeLib.Common.Tools.Logging;
using System;
using System.Reflection;

namespace AccessCodeLib.Common.VBIDETools
{
    public class InvocationHelper : IDisposable
    {
        private object _target;
        private readonly Type _type;

        public InvocationHelper(object target)
        {
            _target = target ?? throw new ArgumentNullException("target");
            _type = _target.GetType();
        }

        public object InvokeMethod(string methodName, object[] args = null)
        {
            return InvokeMember(methodName, BindingFlags.InvokeMethod, args);
        }

        public object InvokePropertyGet(string propertyName)
        {
            return InvokeMember(propertyName, BindingFlags.GetProperty, null);
        }

        public object InvokePropertyGet(string propertyName, object[] args)
        {
            return InvokeMember(propertyName, BindingFlags.GetProperty, args);
        }

        private object InvokeMember(string name, BindingFlags bindingFlags, object[] args)
        {
            var res = _type.InvokeMember(name, bindingFlags, null, _target, args);
            return res;
        }

        #region IDisposable Support

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                _target = null;
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }

            // GC-Bereinigung wegen unmanaged res:
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~InvocationHelper()
        {
            Dispose(false);
        }

        #endregion

    }
}