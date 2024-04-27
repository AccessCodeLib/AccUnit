using System.Windows;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    /// <summary>
    /// Interaktionslogik für ImportExportWindow.xaml
    /// </summary>
    public partial class InsertTestMethodDialog : Window
    {
        public InsertTestMethodDialog(InsertTestMethodViewModel dataContext)
        {
            DataContext = dataContext;
            InitializeComponent();
        }
    }
}
