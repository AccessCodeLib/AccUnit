using System.Windows;

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
