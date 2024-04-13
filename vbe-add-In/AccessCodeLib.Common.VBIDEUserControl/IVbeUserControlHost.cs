using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AccessCodeLib.Common.VBIDETools
{
    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IVbeUserControlHost
    {
        void HostUserControl(UserControl UserControlToHost);
    }
}
