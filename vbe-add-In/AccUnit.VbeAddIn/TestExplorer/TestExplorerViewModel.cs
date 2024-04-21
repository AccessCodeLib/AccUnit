using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestExplorerViewModel : ITestResultReporter, INotifyPropertyChanged
    {
        private string _selectAllCheckBoxText = Resources.UserControls.SelectListSelectAllCheckboxCaption;
        private string _commitButtonText = Resources.UserControls.SelectListCommitButtonText;

        //public delegate void CommitSelectedTestItemsEventHandler(TestExplorerViewModel sender, CheckableTestItemsEventArgs e);
        //public event CommitSelectedTestItemsEventHandler ItemsSelected;

        public delegate void RefreshTestItemListEventHandler(TestExplorerViewModel sender, CheckableTestItemsEventArgs e);
        public event RefreshTestItemListEventHandler RefreshList;

        public event EventHandler<RunTestsEventArgs> RunTests;
        public event EventHandler CancelTestRun;

        public TestExplorerViewModel()
        {
            _selectAllCheckBoxText = Resources.UserControls.SelectListSelectAllCheckboxCaption;
            _commitButtonText = Resources.UserControls.TestExplorerCommitButtonText;

            TestItems = new TestItems();
            RefreshCommand = new RelayCommand(Refresh);
            CommitCommand = new RelayCommand(Commit);
        }

        private TestItems _testItems;
        public TestItems TestItems
        {
            get => _testItems;
            set
            {
                _testItems = value;
                OnPropertyChanged(nameof(TestItems));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #region ITestResultReporter
        private INotifyingTestResultCollector _testResultCollector;

        public ITestResultCollector TestResultCollector
        {
            get { return _testResultCollector; }
            set
            {
                _testResultCollector = value as INotifyingTestResultCollector;
                InitEventHandler();
            }
        }

        private void InitEventHandler()
        {
            _testResultCollector.TestSuiteStarted += TestResultCollector_TestSuiteStarted;
            _testResultCollector.TestSuiteFinished += TestResultCollector_TestSuiteFinished;
            _testResultCollector.TestSuiteReset += TestResultCollector_TestSuiteReset;
            _testResultCollector.TestTraceMessage += TestResultCollector_TestTraceMessage;
            _testResultCollector.TestFixtureStarted += TestResultCollector_TestFixtureStarted;
            _testResultCollector.TestFixtureFinished += TestResultCollector_TestFixtureFinished;
            _testResultCollector.TestStarted += TestResultCollector_TestStarted;
            _testResultCollector.TestFinished += TestResultCollector_TestFinished;
        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            //TestItems.Clear();

            // remove all items that are not in the current test suite

            foreach (var item in TestItems.ToList())
            {
                if (!testSuite.TestFixtures.Any(tf => tf.Name == item.Name))
                {
                    //TestItems.Remove(item);
                }
            }

            //OnPropertyChanged(nameof(TestItems));

            //ClearLogMessages();
            //_vbeUserControl.Show();
            //LogStringToTextBox("TS started ...");
            //if (testSuite is VBATestSuite vbaTestSuite)
            //    LogStringToTextBox(vbaTestSuite.ActiveVBProject.Name);
        }

        private void TestResultCollector_TestSuiteFinished(ITestSummary summary)
        {
            //LogStringToTextBox(summary.ToString());
            //LogStringToTextBox("TS finished.");
        }

        private void TestResultCollector_TestFixtureStarted(ITestFixture fixture)
        {
            var testItem = FindTestItem(fixture);
            if (testItem != null)
            {
                return;
            }
            TestItems.Add(new TestItem(fixture.Name));
        }

        private void TestResultCollector_TestFixtureFinished(ITestResult result)
        {
            var testItem = FindTestItem(result.Test);
            if (testItem == null)
            {
                return;
            }
            testItem.Result = result.Result.ToString();
            testItem.IsExpanded = !result.Success;
            //if (result.Success)
            //{
            //    testItem.IsChecked = false; 
            //}
            testItem.TestResult = result;
            OnPropertyChanged(nameof(testItem.IsExpanded));
        }

        private void TestResultCollector_TestStarted(ITest test, IgnoreInfo ignoreInfo)
        {
            var testItem = FindTestItem(test);
            if (testItem != null)
            {
                var parent = FindParentTestItem(test);
                if (parent != null)
                {
                    parent.IsExpanded = true;
                }   
                return;
            }

            var parentItem = FindParentTestItem(test);
            if (parentItem == null)
            {
                TestItems.Add(new TestItem(test.Name, true));
            }
            else
            {
                if (test.Parent is IRowTest)
                    parentItem.Children.Add(new TestItem(((IRowTestId)test).RowId));
                else
                    parentItem.Children.Add(new TestItem(test.Name));
                parentItem.IsExpanded = true;
            }
            //OnPropertyChanged(nameof(TestItems));
        }

        private void TestResultCollector_TestFinished(ITestResult result)
        {
            var testItem = FindTestItem(result.Test);
            if (testItem == null)
            {
                return;
            }
            testItem.Result = result.Result.ToString();
            testItem.IsExpanded = !result.Success;
            testItem.TestResult = result;
            if (result.Success)
            {
                testItem.IsChecked = false;
                OnPropertyChanged(nameof(testItem.IsChecked));
            }
            OnPropertyChanged(nameof(testItem.IsExpanded));
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            if (resetmode == ResetMode.RemoveTests)
            {
                TestItems.Clear();
            }
        }

        private void TestResultCollector_TestTraceMessage(string message, ICodeCoverageTracker CodeCoverageTracker)
        {
            //LogStringToTextBox(message);
        }
        #endregion


        private TestItem FindParentTestItem(ITest test)
        {
            var parent = test.Parent as ITestData;
            if (test.Parent == null)
            {
                return null;
            }

            if (test.Parent is ITestFixture testFixture)
            {
                return TestItems.FirstOrDefault(ti => ti.Name == testFixture.Name);
            }

            if (test.Parent is IRowTest rowTest)
            {
                var rowTestFixtureItem = FindParentTestItem(rowTest);  
                return rowTestFixtureItem.Children.FirstOrDefault(ti => ti.Name == rowTest.Name);
            }

            var parentTest = test.Parent as ITest;
            var fixtureItem = FindParentTestItem(parentTest);   
            return fixtureItem.Children.FirstOrDefault(ti => ti.Name == parentTest.Name);
        }

        private TestItem FindTestItem(ITestData testData)
        {
            TestItem parentItem = null;
            ITest test = testData as ITest;
            if (test != null)
            {
                parentItem = FindParentTestItem(test);
            }
            if (parentItem == null)
            {
                return TestItems.FirstOrDefault(ti => ti.Name == testData.Name);
            }

            if (test is IRowTestId rowTestId)
            {
                return parentItem.Children.FirstOrDefault(ti => ti.Name == rowTestId.RowId);
            }   

            return parentItem.Children.FirstOrDefault(ti => ti.Name == testData.Name);
        }


        private bool _selectAll = false;

        public bool SelectAll
        {
            get { return _selectAll; }
            set
            {
                if (_selectAll != value)
                {
                    _selectAll = value;
                    if (TestItems != null)
                    {
                        foreach (var item in TestItems)
                        {
                            item.IsChecked = value;
                        }
                    }
                    OnPropertyChanged(nameof(SelectAll));
                }
            }
        }

        public string SelectAllCheckBoxText
        {
            get => _selectAllCheckBoxText;
            set
            {
                if (_selectAllCheckBoxText != value)
                {
                    _selectAllCheckBoxText = value;
                    OnPropertyChanged(nameof(SelectAllCheckBoxText));
                }
            }
        }

        public string CommitButtonText
        {
            get => _commitButtonText;
            set
            {
                if (_commitButtonText != value)
                {
                    _commitButtonText = value;
                    OnPropertyChanged(nameof(CommitButtonText));
                }
            }
        }

        public ICommand RefreshCommand { get; }
        public ICommand CommitCommand { get; }
        public ImageSource RefreshCommandImageSource
        {
            get
            {
                return UITools.ConvertBitmapToBitmapSource(Resources.Icons.refresh_green);
            }
        }

        protected void Refresh()
        {
            RefreshList?.Invoke(this, new CheckableTestItemsEventArgs(TestItems));
        }

        protected virtual void Commit()
        {
            TestClassList list = new TestClassList();
            list.AddRange(TestItems.Where(ti => ti.IsChecked).Select(ti => ti.TestClassInfo));
            RunTests?.Invoke(this, new RunTestsEventArgs(list));
        }

    }

    public static class TestExplorerInfo
    {
        public const string ProgID = @"AccUnit.VbeAddIn.TestExplorer";
        public const string PositionGuid = @"DB052D8D-8418-4322-ADD9-5DCB8157C8D4";
    }
}
