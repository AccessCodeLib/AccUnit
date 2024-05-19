using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class UserSettingDialog : Form
    {
        public UserSettingDialog()
        {
            InitializeComponent();
            Icon = UITools.ConvertBitmapToIcon(Resources.Icons.settings);
        }

        public object Settings
        {
            get
            {
                return SettingPropertyGrid.SelectedObject; 
            }
            set
            {
                SettingPropertyGrid.SelectedObject = value;
            }
        }

    }
}
