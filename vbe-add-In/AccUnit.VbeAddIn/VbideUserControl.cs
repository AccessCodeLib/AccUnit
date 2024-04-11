using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AccessCodeLib.AccUnit.VbeAddIn.Resources;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static AccessCodeLib.AccUnit.VbeAddIn.MyWindow;

namespace AccessCodeLib.AccUnit.VbeAddIn
{

    public class VbideUserControl<TControl> : IDisposable
            where TControl : class
    {

        private readonly TControl _control;
        private readonly Microsoft.Vbe.Interop.Window _vbeWindow;

        readonly MyWindow myWindow = new MyWindow();
        
        public VbideUserControl(AddIn addIn, string caption, string positionGuid, TControl controlToHost, bool visible = true)
        {
            object docObj = null;
            _vbeWindow = addIn.VBE.Windows.CreateToolWindow(addIn, "AccUnit.VbideUserControlHost",
                                                            caption, positionGuid, ref docObj);

            _vbeWindow.Visible = true;

            if (!(docObj is VbideUserControlHost userControlHost))
            {
                throw new InvalidComObjectException(string.Format("docObj cannot be casted to IVbideUserControlHost"));
            }

            _control = controlToHost;
            if (!(_control is UserControl userControl))
            {
                throw new InvalidComObjectException(string.Format("controlToHost cannot be casted to UserControl"));
            }
            //userControlHost.Visible = true;

            IntPtr hWnd = MyWindow.FindWindow("VBFloatingPalette", caption);
            hWnd = MyWindow.FindWindowEx(hWnd, IntPtr.Zero, null, null);

            if (hWnd == IntPtr.Zero)
                throw new Exception(caption + " hwnd nicht gefunden");

            myWindow.Init(hWnd, userControlHost);
            userControlHost.HostUserControl(userControl);

            if (!visible)
            {
                _vbeWindow.Visible = false;
            }

            /*
            var logger = controlToHost as LoggerControl;
            logger.LogTextBox.AppendText(caption + "\r\n");
            logger.LogTextBox.AppendText(hWnd.ToString() + "\r\n");
            logger.LogTextBox.AppendText(hWnd.ToString("X"));
            */

        }

        public TControl Control { get { return _control; } }
        private Microsoft.Vbe.Interop.Window VbeWindow { get { return _vbeWindow; } }

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
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~VbideUserControl()
        {
            Dispose(false);
        }

        #endregion

    }

    public class MyWindow : NativeWindow
    {

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);


        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetParent(IntPtr hWnd);


        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private const int WM_SIZE = 0x0005;
        private const int WM_WINDOWPOSCHANGED = 0x0047;

        private UserControl _userControl;

        private Size _lastSize;

        public void Init(IntPtr hWnd, UserControl userControl)
        {
            _userControl = userControl;
            base.AssignHandle(hWnd);

            CheckSize();

            // TODO: remove timer
            var timer = new System.Windows.Forms.Timer { Interval = 2000 };
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            CheckSize();
        }

        private void CheckSize()
        {
            GetClientRect(this.Handle, out RECT rect);
            var newSize = new Size(rect.Right - rect.Left, rect.Bottom - rect.Top);
            if (newSize != _lastSize)
            {
                // Die Größe des Fensters hat sich geändert
                _userControl.Width = newSize.Width;
                _userControl.Height = newSize.Height;
                _lastSize = newSize;
            }
        }

        /*
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_WINDOWPOSCHANGED)
            {
                GetClientRect(this.Handle, out RECT rect);
                _userControl.Width = rect.Right - rect.Left;
                _userControl.Height = rect.Bottom - rect.Top;
                throw new Exception("xx passt");
            }

            throw new Exception("xxxx");

            base.WndProc(ref m);
        }
        */
    }

    


}
