using System.Windows;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class GenerateTestMethodsFromCodeModuleDialog : Window
    {
        public GenerateTestMethodsFromCodeModuleDialog(GenerateTestMethodsFromCodeModuleViewModel dataContext)
        {
            DataContext = dataContext;
            InitializeComponent();
        }
    }
}
