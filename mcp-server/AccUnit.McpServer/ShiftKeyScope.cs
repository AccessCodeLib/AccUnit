using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Holds the Shift key down until disposed. Access skips AutoExec and the startup form
    /// when Shift is down while a database opens; there is no automation switch for this.
    /// The key state is system-wide, so keep the scope as short as possible.
    /// </summary>
    internal sealed class ShiftKeyScope : IDisposable
    {
        private const byte VirtualKeyShift = 0x10;
        private const uint KeyEventKeyUp = 0x0002;

        private bool _released;

        public ShiftKeyScope()
        {
            keybd_event(VirtualKeyShift, 0, 0, UIntPtr.Zero);
        }

        public void Dispose()
        {
            if (_released) return;
            keybd_event(VirtualKeyShift, 0, KeyEventKeyUp, UIntPtr.Zero);
            _released = true;
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);
    }
}
