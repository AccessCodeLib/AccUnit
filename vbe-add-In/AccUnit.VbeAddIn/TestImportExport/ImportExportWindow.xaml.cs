using System.Windows;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    /// <summary>
    /// Interaktionslogik für ImportExportWindow.xaml
    /// </summary>
    public partial class ImportExportWindow : Window
    {
        public ImportExportWindow(SelectControlViewModel dataContext)
        {
            DataContext = dataContext;
            InitializeComponent();
        }
    }
}
