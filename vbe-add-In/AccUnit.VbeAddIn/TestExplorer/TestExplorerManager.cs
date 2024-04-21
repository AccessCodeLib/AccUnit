using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    internal class TestExplorerManager : ITestResultReporter, ICommandBarsAdapterClient
    {
        private readonly VbeUserControl<TestExplorerView> _vbeUserControl;
        private readonly TestExplorerViewModel _viewModel;
        private INotifyingTestResultCollector _testResultCollector;

        public event EventHandler<RunTestsEventArgs> RunTests;
        public event EventHandler CancelTestRun;

        public TestExplorerManager(VbeUserControl<TestExplorerView> vbeUserControl)
        {
            _vbeUserControl = vbeUserControl;
            _viewModel = new TestExplorerViewModel();
            _vbeUserControl.Control.DataContext = _viewModel;

            InitViewModel();
        }

        private void InitViewModel()
        {
            _viewModel.RefreshList += (sender, e) =>
            {
                FillTestsFromVbProject();
            };  

            _viewModel.RunTests += (sender, e) =>
            {
                RunTests?.Invoke(sender, e);
            };  

            _viewModel.CancelTestRun += (sender, e) =>
            {
                CancelTestRun?.Invoke(sender, e);
            };

            _viewModel.GetTestClassInfo += (sender, e) =>
            {
                e.TestClassInfo = VbeIntegrationManager.TestClassManager.GetTestClassInfo(e.ClassName, true);
            };  
        }   

        public VbeIntegrationManager VbeIntegrationManager { get; set; }

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

 #region ICommandBarsAdapterClient support

        public void SubscribeToCommandBarAdapter(VbeCommandBarAdapter commandBarAdapter)
        {
            using (new BlockLogger())
            {
                var accUnitCommandBarAdapter = commandBarAdapter as AccUnitCommandBarAdapter;
                if (accUnitCommandBarAdapter != null)
                {
                    AddTestListCommandBarButton(commandBarAdapter, accUnitCommandBarAdapter.AccUnitCommandbar);
                }

                // TODO: Why shouldn't there be any view commandbar?
                var viewCommandBar = GetViewCommandBarOrNull(commandBarAdapter);
                if (viewCommandBar != null)
                {
                    const int projectExplorerControlID = 2557;
                    var projectExplorerControlIndex = VbeCommandBarAdapter.GetButtonIndex(viewCommandBar, projectExplorerControlID);
                    AddTestListCommandBarButton(commandBarAdapter, viewCommandBar, projectExplorerControlIndex);
                }
                else
                {
                    if (accUnitCommandBarAdapter != null)
                    {
                        var accUnitSubMenu = accUnitCommandBarAdapter.AccUnitSubMenu.CommandBar;
                        AddTestListCommandBarButton(commandBarAdapter, accUnitSubMenu);
                    }
                }
            }
        }

        private static CommandBar GetViewCommandBarOrNull(VbeCommandBarAdapter commandBarAdapter)
        {
            try
            {
                return commandBarAdapter.CommandBarView;
            }
            catch
            {
                return null;
            }
        }

        private void AddTestListCommandBarButton(VbeCommandBarAdapter commandBarAdapter, CommandBar commandBar, int? index = null)
        {
            var buttonData = new CommandbarButtonData
            {
                Caption = Resources.VbeCommandbars.ViewTestExplorerCommandbarButtonCaption,
                Description = Resources.VbeCommandbars.SelectTestsCommandBarButtonDescription,
                FaceId = 2529,
                BeginGroup = true,
                Index = index
            };
            var button = commandBarAdapter.AddCommandBarButton(commandBar, buttonData, AccUnitCommandBarItemsShowTestListWindow);
            button.Style = MsoButtonStyle.msoButtonAutomatic;
        }

        void AccUnitCommandBarItemsShowTestListWindow(CommandBarButton ctrl, ref bool cancelDefault)
        {
            _vbeUserControl.Visible = true;
            if (_viewModel.TestItems.Count == 0)
            {
                FillTestsFromVbProject();
            }
        }

        private void FillTestsFromVbProject()
        {
            _viewModel.TestItems.Clear();
            var testItems = VbeIntegrationManager.TestClassManager.GetTestClassListFromVBProject(true);   
            foreach (var testItem in testItems)
            {
                _viewModel.TestItems.Add(new TestClassInfoTestItem(testItem, false));
            }
        }

        #endregion

    }
}
