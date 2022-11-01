using System.Runtime.InteropServices;
using AccessCodeLib.AccUnit.Common;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("BA74B9B5-0E96-4648-AC25-78B8052845BC")]
    public interface ITestMessageBox
    {
        int Show(object Prompt, int Buttons = 1, object Title = null, 
                 object HelpFile = null, object Context = null);

        MessageBoxData[] MessageBoxData { get; }

        void InitMsgBoxResults(params VbMsgBoxResult[] args);

        void ActivateTestMessageBox(OfficeApplicationHelper officeApplicationHelper, ITestMessageBox messageBox);
    }

    [ComVisible(true)]
    public enum VbMsgBoxResult
    {
        vbAbort = VBA.VbMsgBoxResult.vbAbort,
        vbCancel = VBA.VbMsgBoxResult.vbCancel,
        vbIgnore = VBA.VbMsgBoxResult.vbIgnore,
        vbNo = VBA.VbMsgBoxResult.vbNo,
        vbOK = VBA.VbMsgBoxResult.vbOK,
        vbRetry = VBA.VbMsgBoxResult.vbRetry,
        vbYes = VBA.VbMsgBoxResult.vbYes
    }
}