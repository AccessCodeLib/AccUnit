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
        private INotifyingTestResultCollector _testResultCollector; 
        
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
                _testResultCollector = value as INotifyingTestResultCollector;
                _testResultCollector.TestSuiteStarted += TestResultCollector_TestSuiteStarted;  
            }
        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            _vbeUserControl.Visible = true; 
        }
    }
}
