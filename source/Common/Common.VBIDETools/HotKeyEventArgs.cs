using System;

namespace AccessCodeLib.Common.VBIDETools
{
    public class HotKeyEventArgs : EventArgs
    {
        public HotKeyEventArgs(HotKey.ModKeys modifiers, uint key)
        {
            Modifiers = modifiers;
            Key = key;
        }

        public HotKey.ModKeys Modifiers { get; private set; }
        public uint Key { get; private set; }
    }
}