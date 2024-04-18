using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using System.Windows;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Documents;
using System;

namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestExplorerViewModel : ITestResultReporter, INotifyPropertyChanged
    {
        public TestExplorerViewModel()
        {
            TestItems = new TestItems();
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
            TestItems.Clear();
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
            TestItems.Add(new TestItem { Name = fixture.Name });
                //OnPropertyChanged(nameof(TestItems));
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
            testItem.TestResult = result;
            //OnPropertyChanged(nameof(TestItems));
            OnPropertyChanged(nameof(testItem.IsExpanded));
        }

        private void TestResultCollector_TestStarted(ITest test, IgnoreInfo ignoreInfo)
        {
            var parentItem = FindParentTestItem(test);
            if (parentItem == null)
            {
                TestItems.Add(new TestItem { Name = test.Name, IsExpanded = true });
            }
            else
            {
                if (test.Parent is IRowTest)
                    parentItem.Children.Add(new TestItem { Name = ((IRowTestId)test).RowId });
                else
                    parentItem.Children.Add(new TestItem { Name = test.Name });
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
            OnPropertyChanged(nameof(testItem.IsExpanded));
        }

        private void TestResultCollector_TestSuiteReset(ResetMode resetmode, ref bool cancel)
        {
            TestItems.Clear();
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
    }

    public static class TestExplorerInfo
    {
        public const string ProgID = @"AccUnit.VbeAddIn.TestExplorer";
        public const string PositionGuid = @"DB052D8D-8418-4322-ADD9-5DCB8157C8D4";
    }
}
