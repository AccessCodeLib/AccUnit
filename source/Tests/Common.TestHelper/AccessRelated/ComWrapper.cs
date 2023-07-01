using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Common.TestHelpers.AccessRelated
{
    public class ComWrapper<T> : IDisposable where T : class
    {
        private readonly T _comReference;

        public ComWrapper(T comReference)
        {
            _comReference = comReference ?? throw new ArgumentNullException("comReference");
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