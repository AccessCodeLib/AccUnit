using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Common
{
    [ComVisible(true)]
    [Guid("D4E58173-675A-49DA-A3AF-AEF4747DC812")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgIdAttribute("AccUnit.MessageBoxData")]
    public class MessageBoxData
    {
        public object Prompt;
        public int Buttons;
        public object Title;
        public object HelpFile;
        public object Context;
        public int MsgBoxResult;
    }
}