using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class LoggerControl : UserControl
    {
        public LoggerControl()
        {
            InitializeComponent();
        }

        public TextBox LogTextBox
        {
            get { return logTextBox; }
        }
    }

    public static class LoggerControlInfo
    {
        public const string ProgID = @"AccUnit.VbeAddIn.LoggerControl";
        public const string PositionGuid = @"68D8D91F-29D9-4672-837F-B4D2BBA730C9";
    }

}
