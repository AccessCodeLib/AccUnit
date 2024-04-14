using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    /// <summary>
    /// Interaktionslogik für TestExplorerTreeView.xaml
    /// </summary>
    public partial class TestExplorerTreeView : UserControl
    {
        public ObservableCollection<TestItem> TestItems { get; set; }

        public TestExplorerTreeView()
        {
            InitializeComponent();
            DataContext = new TestExplorerViewModel();
        }
    }
}
