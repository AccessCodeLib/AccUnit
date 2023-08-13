﻿using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.AccUnit.TestRunner;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AccessCodeLib.AccUnit
{
    public class VBATestSuite : IVBATestSuite, IDisposable, ITestData
    {
        public VBATestSuite()
        {
            using (new BlockLogger())
            {
                SummaryFormatter = new TestSummaryFormatter(TestSuiteUserSettings.Current.SeparatorMaxLength, TestSuiteUserSettings.Current.SeparatorChar);
                _testBuilder.OfficeApplicationReferenceRequired += OnOfficeApplicationReferenceRequired;
            }
        }

        private readonly List<ITestManagerBridge> _accUnitTests = new List<ITestManagerBridge>();
        private readonly List<ITestFixture> _testFixtures = new List<ITestFixture>();
        private IEnumerable<ITestItemTag> _filterTags = null;

        public IEnumerable<ITestFixture> TestFixtures { get { return _testFixtures; } }

        private ITestSummary _testSummary;
        private TestSummaryFormatter SummaryFormatter { get; set; }
        private readonly VBATestBuilder _testBuilder = new VBATestBuilder();

        private ITestRunner _testRunner;
        public ITestResultCollector TestResultCollector { get; set; }

        #region TestSuite Events

        private void OnTestSuiteStarted(ITestSuite testSuite, ITagList tags)
        {
            using (new BlockLogger(testSuite.Name))
            {
                RaiseTestSuiteStarted(testSuite, tags);
            }
        }

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

        void OnTestSuiteFinished(ITestResult result)
        {
            if (Cancel) return;
            using (new BlockLogger(result.Message))
            {
                RaiseTraceMessage(SummaryFormatter.GetTestSuiteFinishedText(result));
                RaiseTestSuiteFinished(result);
            }
        }

        private void OnTestSuiteTestFixtureFinished(ITestResult result)
        {
            if (Cancel) return;
            using (new BlockLogger(result.Message))
            {
                RaiseTraceMessage(SummaryFormatter.GetTestFixtureFinishedText(result));
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
                RaiseTraceMessage(SummaryFormatter.GetTestFixtureStartedText(fixture));
                RaiseTestFixtureStarted(fixture);
            }
        }

        public bool Cancel { get; set; }

        void OnTestSuiteTestStarted(ITest test, IgnoreInfo ignoreInfo, ITagList tags)
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
                    ignoreInfo = memberinfo.IgnoreInfo;
                    ignoreMember = ignoreInfo.Ignore;
                    if (ignoreMember)
                    {
                        SetRunstateToIgnored(test);
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

                RaiseTestStarted(test, ignoreInfo, memberinfo?.Tags);
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
                RaiseTraceMessage(SummaryFormatter.GetTestCaseFinishedText(result));
                // TODO: Here, a TestConverter comes along, which does not implement ITestCase, so the following condition always evaluates to false!
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
                        TestMessageBox.DisposeTestMessageBox(_testBuilder.OfficeApplicationHelper);
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

        private void RaiseTestSuiteStarted(ITestSuite testSuite, ITagList tags)
        {
            TestSuiteStarted?.Invoke(testSuite, tags);
        }

        private void RaiseTestSuiteFinished(ITestResult result)
        {
            TestSuiteFinished?.Invoke(result);
        }

        private void RaiseTestFixtureFinished(ITestResult result)
        {
            TestFixtureFinished?.Invoke(result);
        }

        private void RaiseTestFixtureStarted(ITestFixture fixture)
        {
            TestFixtureStarted?.Invoke(fixture);
        }

        private void RaiseTestStarted(ITest testcase, IgnoreInfo ignoreInfo, ITagList tags)
        {
            TestStarted?.Invoke(testcase, ignoreInfo, tags);
        }

        private void RaiseTestFinished(ITestResult result)
        {
            TestFinished?.Invoke(result);
        }

        private void RaiseTraceMessage(string text)
        {
            TestTraceMessage?.Invoke(text);
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
                if (_testRunner is null)
                {
                    _testRunner = new VbaTestRunner(_testBuilder.ActiveVBProject);
                }
                return _testRunner;
            }
            set
            {
                if (_testRunner != null)
                    try
                    {
                        _testRunner.TestStarted -= OnTestSuiteTestStarted;
                        _testRunner.TestFinished -= OnTestSuiteTestFinished;
                        _testRunner.TestFixtureFinished -= OnTestSuiteTestFixtureFinished;
                        _testRunner.TestFixtureStarted -= OnTestSuiteTestFixtureStarted;
                        _testRunner.TestSuiteStarted -= OnTestSuiteStarted;
                        _testRunner.TestSuiteFinished -= OnTestSuiteFinished;
                    }
                    catch (Exception ex) { Logger.Log(ex); }

                _testRunner = value;
                if (_testRunner != null)
                {
                    _testRunner.TestStarted += OnTestSuiteTestStarted;
                    _testRunner.TestFinished += OnTestSuiteTestFinished;
                    _testRunner.TestFixtureFinished += OnTestSuiteTestFixtureFinished;
                    _testRunner.TestFixtureStarted += OnTestSuiteTestFixtureStarted;
                    _testRunner.TestSuiteStarted += OnTestSuiteStarted;
                    _testRunner.TestSuiteFinished += OnTestSuiteFinished;
                }
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
            fixture.FillFixtureTags(_testBuilder.ActiveVBProject);
            fixture.FillInstanceMembers(_testBuilder.ActiveVBProject);
            fixture.FillTestListFromTestClassInstance(_testBuilder.ActiveVBProject);
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
            Reset(ResetMode.RemoveTests);
            AddToTestSuite(_testBuilder.CreateTestsFromVBProject());
            return this;
        }

        public ITestSuite Filter(IEnumerable<ITestItemTag> filterTags)
        {
            _filterTags = new List<ITestItemTag>(filterTags);
            return this;
        }

        private object _hostApplication;
        public virtual object HostApplication
        {
            get { return _hostApplication; }
            set
            {
                _hostApplication = value;
                _testBuilder.HostApplication = _hostApplication;
            }
        }

        private void OnOfficeApplicationReferenceRequired(ref object returnedObject)
        {
            returnedObject = HostApplication;
        }

        ITestSuite ITestSuite.Reset(ResetMode mode) { return Reset(mode); }

        public virtual IVBATestSuite Reset(ResetMode mode = ResetMode.ResetTestData)
        {
            if (TestSuiteReset != null)
            {
                var cancel = false;
                TestSuiteReset(mode, ref cancel);
                if (cancel)
                    return this;
            }

            _testSummary?.Reset();

            if (TestResultCollector is ITestSummaryTestResultCollector testSummaryCollector)
                testSummaryCollector.Summary.Reset();

            //ConstantsReader.Clear();
            _accUnitTests.Clear();

            // clear Memberinfo (maybe source code changed)
            _testCaseInfos.Clear();

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
                TestResultCollector = new TestResultCollection(this);
            }
            var testResult = TestRunner.Run(this, TestResultCollector, null, _filterTags);
            _testSummary = testResult as ITestSummary;

            RaiseTraceMessage(SummaryFormatter.GetTestSummaryText(Summary));
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
        public event FinishedEventHandler TestSuiteFinished;
        public event TestFixtureStartedEventHandler TestFixtureStarted;
        public event TestStartedEventHandler TestStarted;
        public event FinishedEventHandler TestFinished;
        public event FinishedEventHandler TestFixtureFinished;
        public event MessageEventHandler TestTraceMessage;
        public event TestSuiteResetEventHandler TestSuiteReset;

        #endregion

        private VBProject _activeVbProject;
        public virtual VBProject ActiveVBProject
        {
            get
            {
                return _activeVbProject;
            }
            set
            {
                _activeVbProject = value;
                _testBuilder.ActiveVBProject = _activeVbProject;
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
        }

        void DisposeUnmanagedResources()
        {
            _activeVbProject = null;
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
