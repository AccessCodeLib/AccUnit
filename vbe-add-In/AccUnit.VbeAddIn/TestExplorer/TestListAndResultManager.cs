/*
using System;
using AccessCodeLib.AccUnit.Common;
using AccessCodeLib.AccUnit.Common.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools.Commandbar;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class TestListAndResultManager : IDisposable, ICommandBarsAdapterClient
    {
        private TestSuiteManager _testSuiteManager;
        //private TestListAndResultView _testListAndResultView;
        //private VbideUserControl<TestListAndResultView> _testListAndResultVbideControl;
        //private TagListManager _tagListManager;

        public TestSuiteManager TestSuiteManager
        {
            get { return _testSuiteManager; }
            set
            {
                _testSuiteManager = value;
                _testSuiteManager.TestSuiteStarted += TestSuiteManagerTestSuiteStarted;
                _testSuiteManager.TestStarted += TestSuiteManagerTestStarted;
                _testSuiteManager.TestFinished += TestSuiteManagerTestFinished;
                _testSuiteManager.TestCountChanged += TestSuiteManagerTestCountChanged;
            }
        }

        public event EventHandler<RunTestsEventArgs> RunTests;
        public event EventHandler CancelTestRun;

        public Microsoft.Vbe.Interop.AddIn AddIn { get; set; }
        public VbeIntegrationManager VbeIntegrationManager { get; set; }
        public TestClassManager TestClassManager { get { return VbeIntegrationManager.TestClassManager; } }

        void TestSuiteManagerTestSuiteStarted(ITestSuite testSuite)
        {
            TestListAndResultWindow.Visible = true;
        }

        void TestSuiteManagerTestStarted(ITest test, bool disableTestCaseSelection, bool newTestRun, IgnoreInfo ignoreInfo, TagList tags)
        {
            TestListAndResultWindow.Control.Add(test, disableTestCaseSelection, newTestRun, ignoreInfo, tags);
        }

        void TestSuiteManagerTestFinished(ITestResult result, bool isSummary = false, TestClassMemberInfo memberinfo = null)
        {
            using (new BlockLogger())
            {
                Logger.Log($"Result:{result.Message}");
                TestListAndResultWindow.Control.Add(result, isSummary, memberinfo);
            }
            
        }

        void TestSuiteManagerTestCountChanged(int number)
        {
            TestCount += number;
        }

        public TagListManager TagListManager 
        { 
            get { return _tagListManager; }
            set 
            { 
                _tagListManager = value;
                _tagListManager.TagsSelected += TagsSelected;
            }
        }

        public VbideUserControl<TestListAndResultView> TestListAndResultWindow
        {
            get
            {
                if (_testListAndResultVbideControl == null)
                {
                    _testListAndResultVbideControl = new VbideUserControl<TestListAndResultView>(AddIn, Resources.UserControls.TestResultUserControlInfoCaption, TestListAndResultViewUserControlInfo.PositionGuid);
                    _testListAndResultView = _testListAndResultVbideControl.Control;
                    _testListAndResultView.RunBreakOnErrorMenuItemEnabled = (TestSuiteManager.TestSuite is AccessTestSuite);

                    _testListAndResultView.RunTests += TestListAndResultViewRunTests;
                    _testListAndResultView.ShowSourceCodeInvoked += TestListAndResultViewShowSourceCodeInvoked;
                    _testListAndResultView.ShowTestResultDetailInvoked += TestListAndResultViewShowTestResultDetailInvoked;
                    _testListAndResultView.RefreshTestList += TestListAndResultViewRefreshTestList;
                    _testListAndResultView.Cancel += TestListAndResultViewCancel;
                }
                return _testListAndResultVbideControl;
            }
        }

        void TestListAndResultViewCancel(object sender, EventArgs e)
        {
            if (CancelTestRun != null)
                CancelTestRun(sender, e);
        }

        void TestListAndResultViewRunTests(object sender, RunTestsEventArgs e)
        {
            if (RunTests != null)
                RunTests(sender, e);
        }

        public void AddTestClassListToTestListAndResultWindow()
        {
            try
            {
                _testListAndResultView.Add(TestClassManager.GetTestClassListFromVBProject(TagListManager.GetFilterTagList()));
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        public int TestCount
        {
            get { return _testListAndResultView.TestCount; }
            set { _testListAndResultView.TestCount = value; }
        }

        void TestListAndResultViewShowSourceCodeInvoked(object sender, TestNodeEventArgs e)
        {
            ShowSourceCode(e.ClassName, e.MemberName);
        }

        void TestListAndResultViewRefreshTestList(ref TestClassList list)
        {
            list = TestClassManager.GetTestClassListFromVBProject(TagListManager.GetFilterTagList());
        }

        static void TestListAndResultViewShowTestResultDetailInvoked(object sender, TestNodeInfoEventArgs e)
        {
            var testInfo = new TestInfoForm {TestInfo = e.TestNodeInfo};
            testInfo.ShowDialog();
        }

        private void ShowSourceCode(string classname, string membername)
        {
            var codePane = ActivateCodePane(classname, membername);
            EnsureTextCursorIsVisible(codePane);
        }

        private CodePane ActivateCodePane(string classname, string membername = null)
        {
            var modul = TestClassManager.ActiveVBProject.VBComponents.Item(classname).CodeModule;
            var pane = modul.CodePane;
            pane.Show();
            pane.Window.SetFocus();
            var procLine = 1;
            if (!string.IsNullOrEmpty(membername))
            {
                // TODO: Determine upfront if the member does not exist and throw appropriate exception (including name of the missing member)
                procLine = modul.ProcBodyLine[membername, vbext_ProcKind.vbext_pk_Proc];
            }
            pane.SetSelection(procLine, 1, procLine, 1);
            return pane;
        }

        //  VBE bugfix: invisible text cursor in CodePane
        private static void EnsureTextCursorIsVisible(_CodePane codePane)
        {
            var window = codePane.Window;
            window.Visible = false;
            window.Visible = true;
            window.SetFocus();
        }

        void TagsSelected(TagList tags)
        {
            TestListAndResultWindow.Visible = true;
            _testListAndResultView.Add(TestClassManager.GetTestClassListFromVBProject(tags));
        }

        public void ShowTestListWindow(bool visible, bool loadTestListIfEmpty = true)
        {
            using (new BlockLogger())
            {
                try
                {
                    using (new BlockLogger("Try block, visible = " + visible))
                    {
                        TestListAndResultWindow.Visible = visible;
                        if (visible)
                        {
                            if (loadTestListIfEmpty && TestListAndResultWindow.Control.TestCount == 0)
                            {
                                AddTestClassListToTestListAndResultWindow();
                            }
                            TestListAndResultWindow.Control.Repaint(); // issue #71
                        }
                    }
                }
                catch (Exception xcp)
                {
                    if (VbeIntegrationManager != null && VbeIntegrationManager.HostApplication == null)
                    {
                        VbeIntegrationManager.ShowMissingApplicationInfo();
                        return;
                    }
                    UITools.ShowException(xcp);
                }
            }
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
                                     Caption = Resources.VbeCommandbars.ViewTestListCommandbarButtonCaption,
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
            ShowTestListWindow(true);
        }

        #endregion

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeTestListAndResultVbideControl();
                _testSuiteManager = null;
                _tagListManager = null;
            }

            _disposed = true;
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~TestListAndResultManager()
        {
            Dispose(false);
        }

        private void DisposeTestListAndResultVbideControl()
        {
            if (_testListAndResultVbideControl == null)
                return;

            try
            {
                Settings.Default.TestListVisible = _testListAndResultVbideControl.Visible;
                if (_testListAndResultVbideControl.Visible)
                    _testListAndResultVbideControl.Visible = false;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

            if (_testListAndResultView != null)
                try
                {
                    _testListAndResultView.Dispose();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
                finally
                {
                    _testListAndResultView = null;
                    Logger.Log("_testListAndResultView disposed");
                }

            try
            {
                _testListAndResultVbideControl.Dispose();
                _testListAndResultVbideControl = null;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        #endregion
    }
}
*/