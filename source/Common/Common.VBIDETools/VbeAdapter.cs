using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Common.VBIDETools
{
    public class VbeAdapter : IDisposable
    {
        private OfficeApplicationHelper _officeApplicationHelper;

        public event EventHandler MainWindowRButtonDown;

        public OfficeApplicationHelper OfficeApplicationHelper
        {
            get { return _officeApplicationHelper; }
            set
            {
                using (new BlockLogger())
                {
                    _officeApplicationHelper = value;
                    HotKeys = new HotKeys((IntPtr)VBE.MainWindow.HWnd);
                    SubclassVbeMainWindow();
                }
            }
        }
        public VBE VBE { get { return OfficeApplicationHelper.VBE; } }

        public VBProject ActiveVBProject
        {
            get { return OfficeApplicationHelper.CurrentVBProject; }
        }

        public _CodePane ActiveCodePane
        {
            get { return VBE.ActiveCodePane; }
        }

        #region subclassing

        private void SubclassVbeMainWindow()
        {
            using (new BlockLogger())
            {
                try
                {
                    SubclassHWnd((IntPtr)VBE.MainWindow.HWnd);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        private enum WmMessage
        {
            Parentnotify = 0x210,
            RButtondown = 0x204,
            Hotkey = 0x312
        }

        private const int GWL_WNDPROC = -4;

        [DllImport("user32")]
        private static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, Win32WndProc newProc);
        [DllImport("user32")]
        private static extern int CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int wMsg, int wParam, int lParam);

        private delegate int Win32WndProc(IntPtr hWnd, int wMsg, int wParam, int lParam);
        private IntPtr _oldWndProc = IntPtr.Zero;
        private Win32WndProc _newWndProc;

        private void SubclassHWnd(IntPtr hWnd)
        {
            using (new BlockLogger())
            {
                _newWndProc = new Win32WndProc(NewWndProc);
                _oldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, _newWndProc);
            }
        }

        private int NewWndProc(IntPtr hWnd, int wMsg, int wParam, int lParam)
        {
            try
            {
                switch ((WmMessage)wMsg)
                {
                    case WmMessage.Parentnotify:
                        if (wParam == (int)WmMessage.RButtondown && MainWindowRButtonDown != null)
                            MainWindowRButtonDown(this, null);
                        break;
                    case WmMessage.Hotkey:
                        CheckHotKeys(wParam);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

            try
            {
                return CallWindowProc(_oldWndProc, hWnd, wMsg, wParam, lParam);
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return 0;
            }
        }

        public HotKeys HotKeys { get; private set; }

        private void CheckHotKeys(int wParam)
        {
            foreach (var hotKey in HotKeys.Where(hotKey => hotKey.HotKeyAtom == wParam))
            {
                hotKey.RaisePressed();
                return;
            }
        }

        #endregion

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();
            }

            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            _officeApplicationHelper = null;
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~VbeAdapter()
        {
            Dispose(false);
        }

        #endregion

    }

}
