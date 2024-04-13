using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AccessCodeLib.Common.VBIDETools
{
    [ComVisible(true)]
    [Guid(VbeUserControlHostSettings.Guid)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(VbeUserControlHostSettings.ProgId)]
    public partial class VbeUserControlHost : UserControl, IVbeUserControlHost
    {
        private readonly SubClassingResizeWindow _resizeWindow = new SubClassingResizeWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetParent(IntPtr hWnd);

        public VbeUserControlHost()
        {
            InitializeComponent();
        }

        public void HostUserControl(UserControl UserControlToHost)
        {
            _resizeWindow.Init(this, GetParentVbeWindowHandle());
            Controls.Add(UserControlToHost);
            UserControlToHost.Dock = DockStyle.Fill;
        }

        private IntPtr GetParentVbeWindowHandle()
        {
            return GetParent(this.Handle);
        }

        private class SubClassingResizeWindow : NativeWindow
        {
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

            [DllImport("user32.dll", SetLastError = true)]
            private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", SetLastError = true)]
            private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

            [StructLayout(LayoutKind.Sequential)]
            private struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            private Control _userControl;
            private Size _lastSize;

            public void Init(Control userControl, IntPtr ParentHwnd)
            {
                _userControl = userControl;
                base.AssignHandle(ParentHwnd);
                CheckSize();
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
}
