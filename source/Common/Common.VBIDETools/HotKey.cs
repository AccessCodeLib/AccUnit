using AccessCodeLib.Common.Tools.Logging;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Common.VBIDETools
{
    public class HotKey
    {
        [Flags]
        public enum ModKeys
        {
            Alt = 0x1,
            Control = 0x2,
            Shift = 0x4,
            Win = 0x8,
        };

        public event EventHandler<HotKeyEventArgs> Pressed;
        internal void RaisePressed()
        {
            Pressed?.Invoke(this, new HotKeyEventArgs(Modifiers, Key));
        }

        private IntPtr _hwnd;

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern ushort GlobalAddAtom(string lpString);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern ushort GlobalDeleteAtom(ushort atom);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public HotKey(ModKeys modKeys, uint key, IntPtr hOwner)
        {
            Modifiers = modKeys;
            Key = key;
            HotKeyAtom = GlobalAddAtom(string.Format("AccUnitHotKey_{0}_{1}", modKeys, key));
            if (RegisterHotKey(hOwner, HotKeyAtom, (uint)modKeys, key))
                _hwnd = hOwner;
        }

        public ushort HotKeyAtom { get; private set; }
        public ModKeys Modifiers { get; private set; }
        public uint Key { get; private set; }

        ~HotKey()
        {
            try
            {
                if ((int)_hwnd != 0)
                {
                    UnregisterHotKey(_hwnd, HotKeyAtom);
                    _hwnd = (IntPtr)0;
                }
                GlobalDeleteAtom(HotKeyAtom);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

    }
}