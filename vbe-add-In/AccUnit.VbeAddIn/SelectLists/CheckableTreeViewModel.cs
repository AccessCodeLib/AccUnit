using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.AccUnit.VbeAddIn.TestExplorer;
using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableTreeViewModel : INotifyPropertyChanged
    {
        private string _selectAllCheckBoxText = Resources.UserControls.SelectListSelectAllCheckboxCaption;
        private string _commitButtonText = Resources.UserControls.SelectListCommitButtonText;

        public delegate void RefreshTestItemListEventHandler(CheckableTreeViewModel sender, CheckableTestItemsEventArgs e);
        public event RefreshTestItemListEventHandler RefreshList;

        public event EventHandler<RunTestsEventArgs> RunTests;
        //public event EventHandler CancelTestRun;
        public event EventHandler<GetTestClassInfoEventArgs> GetTestClassInfo;

        public CheckableTreeViewModel()
        {
            _selectAllCheckBoxText = Resources.UserControls.SelectListSelectAllCheckboxCaption;
            _commitButtonText = Resources.UserControls.TestExplorerCommitButtonText;

            TestItems = new TestClassInfoTestItems();
            RefreshCommand = new RelayCommand(Refresh);
            CommitCommand = new RelayCommand(Commit);
        }

        private CheckableItems<TestItem> _testItems;
        public CheckableItems<TestItem> TestItems
        {
            get => _testItems;
            set
            {
                _testItems = value;
                TestItems.CollectionChanged += OnChildrenCollectionChanged;
                OnPropertyChanged(nameof(TestItems));
            }
        }

        private void OnChildrenCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is TestItem testItem)
                    {
                        testItem.PropertyChanged += OnChildPropertyChanged;
                    }
                }
            }
        }

        private void OnChildPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(TestItem.IsChecked)) 
            {
                if (sender is TestItem testItem)
                {
                    if (!testItem.IsChecked)
                    {
                        if (SelectAll == true)
                        {
                            _selectAll = false;
                            OnPropertyChanged(nameof(SelectAll));
                        }
                    }
                    else if (SelectAll == false)
                    {
                        foreach (var item in TestItems)
                        {
                            if (!item.IsChecked)
                                return;
                        }
                        _selectAll = true;
                        OnPropertyChanged(nameof(SelectAll));
                    }
                }
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
            _testResultCollector.TestFixtureStarted += TestResultCollector_TestFixtureStarted;
            _testResultCollector.TestFixtureFinished += TestResultCollector_TestFixtureFinished;
            _testResultCollector.TestStarted += TestResultCollector_TestStarted;
            _testResultCollector.TestFinished += TestResultCollector_TestFinished;
        }

        private void TestResultCollector_TestSuiteStarted(ITestSuite testSuite)
        {
            // remove all items that are not in the current test suite
            foreach (var item in TestItems.ToList())
            {
                if (!testSuite.TestFixtures.Any(tf => tf.FullName == item.FullName))
                {
                    TestItems.Remove(item);
                }
            }
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
            AddTestClassInfoTestItem(fixture.Name);
        }

        private void AddTestClassInfoTestItem(string className)
        {
            var testClassInfoEventArgs = new GetTestClassInfoEventArgs(className);
            GetTestClassInfo?.Invoke(this, testClassInfoEventArgs);
            if (testClassInfoEventArgs.TestClassInfo == null)
            {
                throw new Exception("Test class info not found for test fixture " + className);
            }
            TestItems.Add(new TestClassInfoTestItem(testClassInfoEventArgs.TestClassInfo,true));
        }

        private void TestResultCollector_TestFixtureFinished(ITestResult result)
        {

            var testItem = FindTestItem(result.Test);
            if (testItem == null)
            {
                return;
            }
            testItem.Result = result.Result.ToString();
            testItem.IsExpanded = result.IsFailure || result.IsError;
            testItem.TestResult = result;
            OnPropertyChanged(nameof(testItem.IsExpanded));
        }

        private void TestResultCollector_TestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            var testItem = FindTestItem(test);
            if (testItem != null)
            {
                if (!testItem.IsChecked)
                {
                    ignoreInfo.Ignore = true;
                    ignoreInfo.Comment = "Test is not selected";
                    return;
                }

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
                AddTestClassInfoTestItem(test.Name);
            }
            else
            {
                TestItem child;
                if (test.Parent is IRowTest)
                    child = new TestItem(test.FullName, ((IRowTestId)test).RowId, parentItem.IsChecked);
                else
                    child = new TestItem(test.FullName, test.Name, parentItem.IsChecked);
                child.IsChecked = true;
                parentItem.Children.Add(child);
                parentItem.IsExpanded = true;
            }
        }

        private void TestResultCollector_TestFinished(ITestResult result)
        {
            var testItem = FindTestItem(result.Test);
            if (testItem == null)
            {
                return;
            }

            //testItem.Result = result.Result + result.Message ?? string.Empty;
            testItem.IsExpanded = result.IsFailure || result.IsError;
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
            if ((resetmode & ResetMode.ResetTestSuite) == ResetMode.ResetTestSuite)
            {
                TestItems.Clear();
            }   
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
                return TestItems.FirstOrDefault(ti => ti.FullName == testFixture.FullName);
            }

            if (test.Parent is IRowTest rowTest)
            {
                var rowTestFixtureItem = FindParentTestItem(rowTest);  
                return rowTestFixtureItem.Children.FirstOrDefault(ti => ti.FullName == rowTest.FullName);
            }

            var parentTest = test.Parent as ITest;
            var fixtureItem = FindParentTestItem(parentTest);   
            return fixtureItem.Children.FirstOrDefault(ti => ti.FullName == parentTest.FullName);
        }

        private TestItem FindTestItem(ITestData testData)
        {
            TestItem parentItem = null;
            if (testData is ITest test)
            {
                parentItem = FindParentTestItem(test);
            }
            if (parentItem == null)
            {
                return TestItems.FirstOrDefault(ti => ti.FullName == testData.FullName);
            }
            return parentItem.Children.FirstOrDefault(ti => ti.FullName == testData.FullName);
        }

        private bool _selectAll = true;
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
                            item.SetChecked(value);
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
            list.AddRange(TestItems.Where(ti => ti.IsChecked).Select(ti => ((TestClassInfoTestItem)ti).TestClassInfo));
            RunTests?.Invoke(this, new RunTestsEventArgs(list));
        }

    }

    public static class TestExplorerInfo
    {
        public const string ProgID = @"AccUnit.VbeAddIn.TestExplorer";
        public const string PositionGuid = @"DB052D8D-8418-4322-ADD9-5DCB8157C8D4";
    }
}
