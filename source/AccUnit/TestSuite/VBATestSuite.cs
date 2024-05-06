using AccessCodeLib.AccUnit.CodeCoverage;
using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AccessCodeLib.AccUnit
{
    public class VBATestSuite : IVBATestSuite, IDisposable, ITestData
    {
        public VBATestSuite(IOfficeApplicationHelper applicationHelper,
                            IVBATestBuilder testBuilder,
                            ITestRunner testRunner,
                            ITestSummaryFormatter testSummaryFormatter)
        {
            using (new BlockLogger())
            {
                _applicationHelper = applicationHelper;
                _testBuilder = testBuilder;
                _summaryFormatter = testSummaryFormatter;
                SetNewTestRunner(testRunner);
            }
        }

        public object Parent { get { return null; } }

        private readonly List<ITestManagerBridge> _accUnitTests = new List<ITestManagerBridge>();
        private readonly List<ITestFixture> _testFixtures = new List<ITestFixture>();
        private IEnumerable<ITestItemTag> _filterTags = null;
        private IEnumerable<string> _methodFilter = null;

        public IEnumerable<ITestFixture> TestFixtures { get { return _testFixtures; } }

        private readonly IOfficeApplicationHelper _applicationHelper;
        protected IOfficeApplicationHelper ApplicationHelper { get { return _applicationHelper; } }

        private ITestSummary _testSummary;
        private readonly ITestSummaryFormatter _summaryFormatter;
        private readonly IVBATestBuilder _testBuilder;
        private ITestRunner _testRunner;

        private ITestResultCollector _testResultCollector;
        public ITestResultCollector TestResultCollector
        {
            get
            {
                if (_testResultCollector is null)
                {
                    _testResultCollector = NewTestResultCollector();
                }
                return _testResultCollector;
            }
            set
            {
                _testResultCollector = value;
            }
        }

        protected virtual ITestResultCollector NewTestResultCollector()
        {
            return new TestResultCollector(this);
        }

        public ICodeCoverageTracker CodeCoverageTracker { get; set; }

        private readonly List<ITestResultReporter> _testResultReportes = new List<ITestResultReporter>();
        public void AppendTestResultReporter(ITestResultReporter reporter)
        {
            reporter.TestResultCollector = TestResultCollector;
            AddTestResultReporterToList(reporter);
        }

        protected virtual void AddTestResultReporterToList(ITestResultReporter reporter)
        {
            _testResultReportes.Add(reporter);
        }

        #region TestSuite Events

        private TestClassMemberInfo GetMemberInfo(string classname, string membername)
        {
            var key = GetTestCaseKey(classname, membername);
            if (!_testCaseInfos.TryGetValue(key, out TestClassMemberInfo memberinfo))
            {
                var reader = new TestClassReader(ActiveVBProject);
                memberinfo = reader.GetTestClassMemberInfo(classname, membername);
                _testCaseInfos.Add(key, memberinfo);
            }
            return memberinfo;
        }

        private void OnTestSuiteTestFixtureFinished(ITestResult result)
        {
            if (Cancel) return;
            using (new BlockLogger(result.Message))
            {
                RaiseTraceMessage(_summaryFormatter.GetTestFixtureFinishedText(result));
                RaiseTestFixtureFinished(result);
            }
        }

        private void OnTestSuiteTestFixtureStarted(ITestFixture fixture)
        {
            if (Cancel)
            {
                fixture.RunState = RunState.Ignored;
                return;
            }
            using (new BlockLogger(fixture.FullName))
            {
                RaiseTestFixtureStarted(fixture);
                RaiseTraceMessage(_summaryFormatter.GetTestFixtureStartedText(fixture));
            }
        }

        public bool Cancel { get; set; }

        void OnTestSuiteTestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            if (Cancel)
            {
                test.RunState = RunState.Ignored;
                return;
            }

            using (new BlockLogger(test.FullName))
            {
                var memberinfo = GetMemberInfo(test);
                var ignoreMember = false;
                if (memberinfo != null)
                {
                    var memberIgnoreInfo = memberinfo.IgnoreInfo;
                    ignoreMember = memberIgnoreInfo.Ignore;
                    if (ignoreMember)
                    {
                        SetRunstateToIgnored(test);
                        ignoreInfo.Ignore = true;
                        ignoreInfo.Comment = memberIgnoreInfo.Comment;
                    }
                }

                if (!ignoreMember)
                {
                    var bridge = TryGetTestManagerBridge(test);
                    if (bridge != null)
                    {
                        var testManager = bridge.GetTestManager();
                        testManager.InitTestMessageBox(test);
                        if (IsRowTest(test))
                        {
                            var row = GetTestRow(testManager, test);
                            if (row != null && row.IgnoreInfo.Ignore)
                            {
                                SetRunstateToIgnored(test);
                                ignoreInfo = row.IgnoreInfo;
                            }
                        }
                    }
                }

                OnTestStarted(memberinfo);
                // HACK ShowAs: Do this in a better way/in a better place
                if (memberinfo != null)
                {
                    // TODO ShowAs: This messes up RowTests
                    //testcase.DisplayName = memberinfo.DisplayName;
                }

                RaiseTestStarted(test, ref ignoreInfo);
            }
        }

        public ITestRow GetTestRow(ITest test)
        {
            var bridge = TryGetTestManagerBridge(test);
            if (bridge != null)
            {
                var testManager = bridge.GetTestManager();
                if (IsRowTest(test))
                {
                    return GetTestRow(testManager, test);
                }
            }

            return null;
        }

        private static ITestRow GetTestRow(TestManager testManager, ITest testCase)
        {
            var member = testManager.Members.Find(
                            m => m.Name.Equals(testCase.Fixture.Name, StringComparison.CurrentCultureIgnoreCase));
            if (member is null)
                return null;

            var row =
                member.TestRows.Find(
                    m => m.TestFixtureRowName.Equals(testCase.Name, StringComparison.CurrentCultureIgnoreCase));

            return row;
        }

        private void SetRunstateToIgnored(ITest test)
        {
            test.RunState = RunState.Ignored;
        }

        private ITestManagerBridge TryGetTestManagerBridge(ITest test)
        {
            try
            {
                var bridge = _accUnitTests.Find(m => m.GetTestManager().TestName == test.Name);
                return bridge;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return null;
            }
        }

        protected virtual void OnTestStarted(TestClassMemberInfo testClassMemberInfo)
        { }

        private readonly IDictionary<string, TestClassMemberInfo> _testCaseInfos = new Dictionary<string, TestClassMemberInfo>();

        private TestClassMemberInfo GetMemberInfo(ITest test)
        {
            Debug.Assert(test != null);
            var classname = test.Fixture.Name;
            return GetMemberInfo(classname, test.MethodName);
        }

        private static bool IsRowTest(ITest test)
        {
            if (test is IRowTest)
            {
                return true;
            }
            return false;
        }

        private static string GetTestCaseKey(string classname, string membername)
        {
            return $"{classname}.{membername}";
        }

        private void OnTestSuiteTestFinished(ITestResult result)
        {
            if (Cancel) return;
            using (new BlockLogger(result.Message))
            {
                if (!(result.Test is IRowTest))
                {
                    RaiseTraceMessage(_summaryFormatter.GetTestCaseFinishedText(result));
                }

                // TODO: Here, a TestConverter comes along, which does not implement ITest, so the following condition always evaluates to false!
                if (result.Test is ITest test)
                {
                    var memberinfo = GetMemberInfo(test);
                    test.DisplayName = memberinfo.DisplayName;
                }
                OnTestFinished(result);
                RaiseTestFinished(result);

                DisposeTestTools(result.Test);
            }
        }

        private void DisposeTestTools(ITestData test)
        {
            using (new BlockLogger(test.Name))
            {
                try
                {
                    if (_testBuilder.TestToolsActivated)
                        TestMessageBox.DisposeTestMessageBox(_applicationHelper);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
        }

        protected virtual void OnTestFinished(ITestResult result) { }

        #endregion

        #region Event-Invocators

        protected virtual void RaiseTestSuiteStarted(ITestSuite testSuite/*, IEnumerable<ITestItemTag> tags*/)
        {
            TestSuiteStarted?.Invoke(testSuite);
        }

        private void RaiseTestSuiteFinished(ITestSummary testSummary)
        {
            TestSuiteFinished?.Invoke(testSummary);
        }

        private void RaiseTestFixtureFinished(ITestResult result)
        {
            TestFixtureFinished?.Invoke(result);
        }

        private void RaiseTestFixtureStarted(ITestFixture fixture)
        {
            TestFixtureStarted?.Invoke(fixture);
        }

        protected virtual void RaiseTestStarted(ITest test, ref IgnoreInfo ignoreInfo)
        {
            TestStarted?.Invoke(test, ref ignoreInfo);
        }

        private void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }

        protected virtual void RaiseTraceMessage(string text)
        {
            TestTraceMessage?.Invoke(text, CodeCoverageTracker);
        }

        #endregion

        #region IVBATestSuite Implementation

        public TestClassMemberInfo GetTestClassMemberInfo(string classname, string membername)
        {
            var key = GetTestCaseKey(classname, membername);
            _testCaseInfos.TryGetValue(key, out TestClassMemberInfo memberinfo);
            return memberinfo;
        }

        public virtual string Name { get { return null; } }
        string ITestData.FullName { get { return Name; } }

        public ITestRunner TestRunner
        {
            get
            {
                return _testRunner;
            }
        }

        private void SetNewTestRunner(ITestRunner testRunner)
        {
            if (_testRunner != null)
            {
                try
                {
                    _testRunner.TestStarted -= OnTestSuiteTestStarted;
                    _testRunner.TestFinished -= OnTestSuiteTestFinished;
                    _testRunner.TestFixtureFinished -= OnTestSuiteTestFixtureFinished;
                    _testRunner.TestFixtureStarted -= OnTestSuiteTestFixtureStarted;
                }
                catch (Exception ex) { Logger.Log(ex); }
            }

            _testRunner = testRunner;
            if (_testRunner != null)
            {
                _testRunner.TestStarted += OnTestSuiteTestStarted;
                _testRunner.TestFinished += OnTestSuiteTestFinished;
                _testRunner.TestFixtureFinished += OnTestSuiteTestFixtureFinished;
                _testRunner.TestFixtureStarted += OnTestSuiteTestFixtureStarted;
            }
        }

        public virtual IVBATestSuite Add(object testToAdd)
        {
            AddToTestSuite(_testBuilder.CreateTest(testToAdd, null));
            return this;
        }

        private void AddToTestSuite(object testToAdd)
        {
            if ((testToAdd as ITestManagerBridge) != null)
                _accUnitTests.Add(testToAdd as ITestManagerBridge);

            var fixture = new TestFixture(testToAdd);
            fixture.FillFixtureTags(_applicationHelper.CurrentVBProject);
            fixture.FillInstanceMembers(_applicationHelper.CurrentVBProject);
            fixture.FillTestListFromTestClassInstance(_applicationHelper.CurrentVBProject);
            _testFixtures.Add(fixture);
        }

        private void AddToTestSuite(IEnumerable<object> testsToAdd)
        {
            foreach (var o in testsToAdd)
            {
                AddToTestSuite(o);
            }
        }

        public virtual void AddTestClasses(IEnumerable<TestClassInfo> testClasses)
        {
            AddToTestSuite(_testBuilder.CreateTests(testClasses));
        }

        public virtual IVBATestSuite AddByClassName(string className)
        {
            AddToTestSuite(_testBuilder.CreateTest(className));
            return this;
        }

        public virtual IVBATestSuite AddFromVBProject()
        {
            RaiseTraceMessage("AddFromVBProject:reset");
            Reset(ResetMode.RemoveTests);
            RaiseTraceMessage("AddFromVBProject:AddToTestSuite");
            AddToTestSuite(_testBuilder.CreateTestsFromVBProject());
            RaiseTraceMessage("AddFromVBProject:Completed");
            return this;
        }

        public ITestSuite Select(IEnumerable<string> methodFilter)
        {
            _methodFilter = new List<string>(methodFilter);
            return this;
        }

        public ITestSuite Filter(IEnumerable<ITestItemTag> filterTags)
        {
            _filterTags = new List<ITestItemTag>(filterTags);
            return this;
        }

        ITestSuite ITestSuite.Reset(ResetMode mode) { return Reset(mode); }

        public virtual IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData)
        {
            /*
                None = 0,
                ResetTestData = 1,
                RemoveTests = 2,
                ResetTestSuite = 4,
                RefreshFactoryModule = 8,
                DeleteFactoryModule = 16
            */

            if (TestSuiteReset != null)
            {
                var cancel = false;
                TestSuiteReset(mode, ref cancel);
                if (cancel)
                    return this;
            }

            //RaiseTraceMessage("Reset: _testSummary");
            _testSummary?.Reset();

            //RaiseTraceMessage("Reset: testSummaryCollector");
            if (TestResultCollector is ITestSummaryTestResultCollector testSummaryCollector)
                testSummaryCollector.Summary.Reset();

            //ConstantsReader.Clear();
            //RaiseTraceMessage("Reset: _accUnitTests");
            _accUnitTests.Clear();

            // clear Memberinfo (maybe source code changed)
            //RaiseTraceMessage("Reset: _testCaseInfos");
            _testCaseInfos.Clear();

            if ((mode & ResetMode.RemoveTests) == ResetMode.RemoveTests)
            {
                //RaiseTraceMessage("Reset: _testFixtures");
                _testFixtures.Clear();
            }

            if ((mode & ResetMode.DeleteFactoryModule) == ResetMode.DeleteFactoryModule)
            {
                _testBuilder.DeleteFactoryCodeModule();
            }

            if ((mode & ResetMode.RefreshFactoryModule) == ResetMode.RefreshFactoryModule)
            {
                _testBuilder.RefreshFactoryCodeModule();
            }
            return this;
        }

        ITestSuite ITestSuite.Run() { return Run(); }

        public virtual IVBATestSuite Run()
        {
            Cancel = false;
            if (TestResultCollector is null)
            {
                TestResultCollector = NewTestResultCollector();
            }
            //RaiseTestSuiteStarted(this, _filterTags);
            RaiseTestSuiteStarted(this);
            var testResult = TestRunner.Run(this, TestResultCollector, _methodFilter, _filterTags);
            _testSummary = testResult as ITestSummary;

            RaiseTraceMessage(_summaryFormatter.GetTestSummaryText(Summary));
            RaiseTestSuiteFinished(Summary);
            return this;
        }

        public virtual ITestSummary Summary
        {
            get
            {
                if (TestResultCollector is ITestSummaryTestResultCollector summaryCollector)
                {
                    return summaryCollector.Summary;
                }
                else
                {
                    return _testSummary;
                }
            }
        }

        public event TestSuiteStartedEventHandler TestSuiteStarted;
        public event TestSuiteFinishedEventHandler TestSuiteFinished;
        public event TestFixtureStartedEventHandler TestFixtureStarted;
        public event TestStartedEventHandler TestStarted;
        public event FinishedEventHandler TestFinished;
        public event FinishedEventHandler TestFixtureFinished;
        public event TestTraceMessageEventHandler TestTraceMessage;
        public event TestSuiteResetEventHandler TestSuiteReset;

        #endregion

        public virtual VBProject ActiveVBProject
        {
            get
            {
                return _applicationHelper.CurrentVBProject;
            }
        }

        #region IDisposable Support

        public event DisposeEventHandler Disposed;

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            try
            {
                if (disposing)
                {
                    DisposeManagedResources();
                }
                DisposeUnmanagedResources();
                _disposed = true;
            }
            catch (Exception ex) { Logger.Log(ex); }

            Disposed?.Invoke(this);

            GC.Collect();
        }

        void DisposeManagedResources()
        {
            _testBuilder.Dispose();
            _applicationHelper.Dispose();
        }

        void DisposeUnmanagedResources()
        {
            _testResultCollector = null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VBATestSuite()
        {
            Dispose(false);
        }

        #endregion

    }
}
