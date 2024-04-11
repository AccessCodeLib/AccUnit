using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    [ComVisible(true)]
    [Guid("030A1F2F-4E0B-4041-A7F5-C4C0B94BAF07")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.VbideUserControlHost")]
    public partial class VbideUserControlHost : UserControl, IVbideUserControlHost
    {
        public VbideUserControlHost()
        {
            InitializeComponent();
        }

        public void HostUserControl(UserControl userControl)
        {
            Controls.Add(userControl);
            userControl.Dock = DockStyle.Fill;
            userControl.Visible = true;
            //userControl.Refresh();
            //this.Refresh();
        }
    }

    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IVbideUserControlHost
    {
        void HostUserControl(UserControl userControl);
    }
}
