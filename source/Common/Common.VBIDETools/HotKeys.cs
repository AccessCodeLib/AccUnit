using System;
using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools
{
    public class HotKeys : List<HotKey>, IDisposable
    {
        private readonly IntPtr _hWnd;
        public HotKeys(IntPtr hWnd)
        {
            _hWnd = hWnd;
        }

        public HotKey RegisterHotKey(HotKey.ModKeys modKeys, uint key)
        {
            var hotKey = new HotKey(modKeys, key, _hWnd);
            Add(hotKey);
            return hotKey;
        }

        /// <summary>
        /// Unregisters all hot keys.
        /// </summary>
        public void Dispose()
        {
            foreach (var hotKey in this)
            {
                hotKey.Dispose();
            }
            Clear();
        }
    }
}