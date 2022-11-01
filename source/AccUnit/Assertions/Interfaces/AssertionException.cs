using System;
using System.IO;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Assertions.Interfaces
{
    [ComVisible(true)]
    [Guid("560707B9-7542-4234-BB3D-0C2632E30098")]
    public interface IAssertionException
    {
        IMatchResult Result { get; }
        string Message { get; }
    }

    [ComVisible(true)]
    [Guid("65C05E6F-959A-4554-8861-1BE212A43BA7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Interop.Constants.ProgIdLibName + ".AssertionException")]
    public class AssertionException : Exception, IAssertionException
    {
        public IMatchResult Result { [return: MarshalAs(UnmanagedType.IDispatch)] get; private set; }
        
        public AssertionException(string message) : base(message)
        {
        }

        public AssertionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public AssertionException(string message, IMatchResult matchResult) : base(message)
        {
            Result = matchResult;
        }
    }
}
