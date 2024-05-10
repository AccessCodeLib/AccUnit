using System.Windows;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class InsertTestMethodDialog : Window
    {
        public InsertTestMethodDialog(InsertTestMethodViewModel dataContext)
        {
            DataContext = dataContext;
            InitializeComponent();
        }
    }
}
