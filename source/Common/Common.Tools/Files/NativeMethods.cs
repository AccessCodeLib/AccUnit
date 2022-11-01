using System;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.Common.Tools.Files
{
    internal class NativeMethods
    {
        [DllImport("mpr.dll", SetLastError = false, CharSet = CharSet.Auto, ThrowOnUnmappableChar = true, BestFitMapping = false)]
        public static extern Int32 WNetGetConnection(string localName, StringBuilder remoteName, ref Int32 remoteSize);
    }
}
