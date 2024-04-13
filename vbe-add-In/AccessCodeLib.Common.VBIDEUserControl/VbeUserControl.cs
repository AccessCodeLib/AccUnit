using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools
{
    public class VbeUserControl<TControl> : IDisposable
            where TControl : UserControl
    {
        private readonly TControl _control;
        private readonly Window _vbeWindow;

        public VbeUserControl(AddIn addIn, string caption, string positionGuid, 
                                TControl controlToHost, bool visible = true,
                                string VbideUserControlHostProgId = VbeUserControlHostSettings.ProgId)
        {
            object docObj = null;
            _vbeWindow = addIn.VBE.Windows.CreateToolWindow(addIn, VbideUserControlHostProgId,
                                                            caption, positionGuid, ref docObj);
            _vbeWindow.Visible = true;

            if (!(docObj is IVbeUserControlHost userControlHost))
            {
                throw new InvalidComObjectException(string.Format("docObj cannot be casted to IVbeUserControlHost"));
            }

            _control = controlToHost;
            userControlHost.HostUserControl(_control);

            if (!visible)
            {
                _vbeWindow.Visible = false;
            }
        }

        public TControl Control { get { return _control; } }
        private Window VbeWindow { get { return _vbeWindow; } }

        public bool Visible
        {
            get
            {
                try
                {
                    return VbeWindow.Visible;
                }
                catch (Exception)
                {
                    return false;
                }
            }
            set { VbeWindow.Visible = value; }
        }

        public void Show()
        {
            if (!Visible)
            {
                Visible = true;
            }   
        }

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
            try
            {
                _vbeWindow.Close();
            }
            catch { /* ignore */ }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VbeUserControl()
        {
            Dispose(false);
        }
        #endregion
    }

    internal class SubClassingResizeWindow : NativeWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private readonly Control _userControl;
        private readonly Microsoft.Vbe.Interop.Window _vbeWindow;
        private Size _lastSize;

        public SubClassingResizeWindow(Control userControl, Microsoft.Vbe.Interop.Window vbeWindow)
        {
            _userControl = userControl;
            _vbeWindow = vbeWindow;

            IntPtr hWnd = FindVbeWindowHostHwnd(_vbeWindow);
            if (hWnd == IntPtr.Zero)
                throw new Exception(string.Concat("hWnd for ", _vbeWindow.Caption, " not found"));

            base.AssignHandle(hWnd);

            CheckSize();
        }

        private static IntPtr FindVbeWindowHostHwnd(Microsoft.Vbe.Interop.Window vbeWindow)
        {
            const string DockedWindowClass = "wndclass_desked_gsk";
            const string FloatingWindowClass = "VBFloatingPalette";
            const string GenericPaneClass = "GenericPane";

            IntPtr hWnd;
            if (IsDockedWindow(vbeWindow))
            {
                hWnd = FindWindow(DockedWindowClass, vbeWindow.LinkedWindowFrame.Caption);
            }
            else
            {
                hWnd = FindWindow(FloatingWindowClass, vbeWindow.Caption);
            }
            hWnd = FindWindowEx(hWnd, IntPtr.Zero, GenericPaneClass, vbeWindow.Caption);
            return hWnd;
        }

        private static bool IsDockedWindow(Microsoft.Vbe.Interop.Window vbeWindow)
        {
            return vbeWindow.LinkedWindowFrame.Type == vbext_WindowType.vbext_wt_MainWindow;
        }

        private void CheckSize()
        {
            GetClientRect(Handle, out RECT rect);
            var newSize = new Size(rect.Right - rect.Left, rect.Bottom - rect.Top);
            if (newSize != _lastSize)
            {
                _userControl.Width = newSize.Width;
                _userControl.Height = newSize.Height;
                _lastSize = newSize;
            }
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_SIZE = 0x0005;

            if (m.Msg == WM_SIZE)
            {
                CheckSize();
            }
            base.WndProc(ref m);
        }
    }
}
