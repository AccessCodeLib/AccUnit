using System.Windows;
using System.Windows.Controls;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public partial class TestExplorerTreeView : UserControl
    {
        public TestExplorerTreeView()
        {
            InitializeComponent();
            //DataContext = new TestExplorerViewModel();
        }

        private void TreeViewItem_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TreeViewItem treeViewItem && treeViewItem.DataContext is TestItem testItem)
            {
                testItem.IsFocused = true;
            }
        }

        private void TreeViewItem_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TreeViewItem treeViewItem && treeViewItem.DataContext is TestItem testItem)
            {
                testItem.IsFocused = false;
            }
        }
    }
}
