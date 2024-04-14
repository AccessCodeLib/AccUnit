using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.VBIDETools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    internal class TestExplorerManager : ITestResultReporter
    {
        private readonly VbeUserControl<TestExplorerTreeView> _vbeUserControl;
        private readonly TestExplorerTreeView _treeView;
        private readonly TestExplorerViewModel _viewModel;
        
        public TestExplorerManager(VbeUserControl<TestExplorerTreeView> vbeUserControl)
        {
            _vbeUserControl = vbeUserControl;
            _treeView = _vbeUserControl.Control;
            _viewModel = _treeView.DataContext as TestExplorerViewModel;
        }

        public ITestResultCollector TestResultCollector
        {
            get { return _viewModel.TestResultCollector; }
            set
            {
                _viewModel.TestResultCollector = value;
            }
        }

    }
}
