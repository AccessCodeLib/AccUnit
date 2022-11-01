using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Common.TestHelpers.AccessRelated
{
    public class ComWrapper<T> : IDisposable where T : class
    {
        private readonly T _comReference;

        public ComWrapper(T comReference)
        {
            if (comReference == null)
                throw new ArgumentNullException("comReference");

            _comReference = comReference;
        }

        public T ComReference
        {
            get { return _comReference; }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_comReference);
        }
    }
}