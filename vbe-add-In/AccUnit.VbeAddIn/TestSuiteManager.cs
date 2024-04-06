using System;
using System.Collections.Generic;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using Application = System.Windows.Forms.Application;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class TestSuiteManager : IDisposable
    {

        public delegate void TestSuiteInitializedEventHandler(ITestSuite suite);
        public event TestSuiteInitializedEventHandler TestSuiteInitialized;

        public delegate void TestSuiteStartedEventHandler(ITestSuite testSuite);
        public event TestSuiteStartedEventHandler TestSuiteStarted;

        public delegate void TestStartedEventHandler(ITest test, bool newTestRun, IgnoreInfo ignoreInfo, IEnumerable<ITestItemTag> tags);
        public event TestStartedEventHandler TestStarted;

        public delegate void TestFinishedEventHandler(ITestResult result, bool isSummary, TestClassMemberInfo memberinfo);
        public event TestFinishedEventHandler TestFinished;

        public delegate void TestCountChangedEventHandler(int number);
        public event TestCountChangedEventHandler TestCountChanged;

        private IVBATestSuite _vbaTestSuite;

        private OfficeApplicationHelper _officeApplicationHelper;
        public OfficeApplicationHelper OfficeApplicationHelper
        {
            get { return _officeApplicationHelper; }
            set 
            { 
                _officeApplicationHelper = value;
                //if (_vbaTestSuite != null)
                //{
                //    ((VBATestSuite)_vbaTestSuite).HostApplication = _officeApplicationHelper.Application;
                //}
            }
        }

        public IVBATestSuite TestSuite
        {
            get
            {
                if (_vbaTestSuite == null)
                {
                    InitTestSuite();
                }
                return _vbaTestSuite;
            }
        }

        public VBProject ActiveVBProject
        {
            get
            {
                return ((VBATestSuite)TestSuite).ActiveVBProject;
            }
        }

        private void InitTestSuite()
        {
            using (new BlockLogger())
            {
                try
                {
                    _vbaTestSuite = CreateVbaTestSuite(OfficeApplicationHelper);
                    if (_vbaTestSuite != null)
                    {
                        _vbaTestSuite.TestSuiteStarted += VBATestsuiteTestSuiteStarted;
                        _vbaTestSuite.TestSuiteFinished += VBATestsuiteTestSuiteFinished;
                        _vbaTestSuite.TestFixtureStarted += VBATestsuiteTestFixtureStarted;
                        _vbaTestSuite.TestFixtureFinished += VBATestsuiteTestFixtureFinished;
                        _vbaTestSuite.TestStarted += VBATestsuiteTestCaseStarted;
                        _vbaTestSuite.TestFinished += VBATestsuiteTestCaseFinished;
                    }
                }
                catch (Exception ex)
                {
                    UITools.ShowException(ex);
                }
                finally
                {
                    if (TestSuiteInitialized != null)
                        TestSuiteInitialized(_vbaTestSuite);
                }
            }
        }

        private static IVBATestSuite CreateVbaTestSuite(OfficeApplicationHelper applicationHelper)
        {
            using (new BlockLogger())
            {
                IVBATestSuite vbaTestSuite;
                var accUnitFactory = new Interop.AccUnitFactory();
                if (applicationHelper is AccessApplicationHelper)
                {
                    Logger.Log("Access application");
                    vbaTestSuite = accUnitFactory.AccessTestSuite(applicationHelper.Application);
                }
                else
                {
                    vbaTestSuite = accUnitFactory.VBATestSuite(applicationHelper.Application);
                }
                return vbaTestSuite;
            }
        }

        void VBATestsuiteTestCaseStarted(ITest test, IgnoreInfo ignoreInfo, IEnumerable<ITestItemTag> tags)
        {
            if (_disableResultOutput)
                return;

            try
            {
                Application.DoEvents();

                Logger.Log(string.Format("ignore?: {0} ... {1}", ignoreInfo.Ignore, ignoreInfo.Comment));

                var parentTest = test.Parent as ITest;
                if (parentTest != null)
                {
                    if (parentTest.RunState == RunState.Ignored)
                    {
                        return;
                    }
                }

                RaiseTestStarted(test, tags, false, ignoreInfo);
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        private bool _disableResultOutput;
        void VBATestsuiteTestCaseFinished(ITestResult result)
        {
            if (_disableResultOutput)
                return;

            try
            {
                using (new BlockLogger())
                {
                    try
                    {
                        Logger.Log(string.Format("Result: {0}", result.Message));

                        if (result.IsError 
                            && !string.IsNullOrEmpty(result.Message) 
                            && result.Message.IndexOf("error 440,", StringComparison.CurrentCultureIgnoreCase)==0)
                        {
                            
                            var accessSuite = _vbaTestSuite as AccessTestSuite;
                            if (accessSuite != null && accessSuite.ErrorTrapping == VbaErrorTrapping.BreakOnAllErrors)
                            {
                                _disableResultOutput = true;
                                return;
                            }   
                        }
                    }
                    catch(Exception ex)
                    {
                        Logger.Log(ex);
                    }
                    /*
                    var test = result.Test.Parent as ITest;
                    if (test != null)
                    {
                        if (test.RunState == RunState.Ignored)
                        {
                            if (TestCountChanged != null)
                                TestCountChanged(-1);
                            return;
                        }
                    }
                    */
                    var memberinfo = GetMemberInfoFromTestCaseResult(result);
                    RaiseTestFinished(result, false, memberinfo);
                }
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        private void VBATestsuiteTestFixtureStarted(ITestFixture fixture)
        {
            if (_disableResultOutput)
                return;
            
            try
            {
                var codeModule = GetCodeModule(fixture.Name);
                var headerText = GetHeaderText(codeModule);
                var tags = GetTagsFromTestClass(fixture.Name, headerText);
                Logger.Log("VBATestsuiteTestFixtureStarted:" + fixture.Name);
                RaiseTestStarted((ITest)fixture, tags);
            }
            catch (Exception xcp)
            {
                Logger.Log(fixture.Name);
                UITools.ShowException(xcp);
            }
        }

        private ITagList GetTagsFromTestClass(string name, string headerText)
        {
            return new TestClassInfo(name, headerText, null).Tags;
        }

        public _CodeModule GetCodeModule(string name)
        {
            using (new BlockLogger())
            {
                return ActiveVBProject.VBComponents.Item(name).CodeModule;
            }
        }

        private static string GetHeaderText(_CodeModule codeModule)
        {
            return codeModule.Lines[1, codeModule.CountOfDeclarationLines];
        }

        void VBATestsuiteTestFixtureFinished(ITestResult result)
        {
            if (_disableResultOutput)
                return;

            try
            {
                RaiseTestFinished(result);
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        void VBATestsuiteTestSuiteStarted(ITestSuite testSuite, IEnumerable<ITestItemTag> tags)
        {
            _disableResultOutput = false;

            try
            {
                var test = (ITest)testSuite;
                var newTestRun = (test.Parent == null);
                if (newTestRun && TestSuiteStarted != null)
                {
                    TestSuiteStarted(testSuite);
                }
                RaiseTestStarted(test, tags, newTestRun);
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        void VBATestsuiteTestSuiteFinished(ITestResult result)
        {
            if (_disableResultOutput)
                return;

            try
            {
                var test = result.Test as ITest;
                var parent = test?.Parent as ITest;
                var isSummary = (parent == null);
                TestClassMemberInfo memberinfo = null;
                if (!isSummary)
                {
                    memberinfo = _vbaTestSuite.GetTestClassMemberInfo(parent.Name, test.Name);
                }
                RaiseTestFinished(result, isSummary, memberinfo);
            }
            catch (Exception xcp)
            {
                UITools.ShowException(xcp);
            }
        }

        private void RaiseTestStarted(ITest test, IEnumerable<ITestItemTag> tags, bool newTestRun = false, IgnoreInfo ignoreInfo = null)
        {
            if (TestStarted != null)
                TestStarted(test, newTestRun, ignoreInfo, tags);
        }

        private void RaiseTestFinished(ITestResult result, bool isSummary = false, TestClassMemberInfo memberinfo = null)
        {
            if (TestFinished != null)
                TestFinished(result, isSummary, memberinfo);
        }

        TestClassMemberInfo GetMemberInfoFromTestCaseResult(ITestResult result)
        {
            var test = result.Test;

            if (!(test?.Parent is ITestData parent))
            {
                throw new NullReferenceException("result.Test.Parent");
            }

            /*
            if (test.IsSuite)
            {
                test = parent;
                parent = (test != null) ? test.Parent as ITest : null;
            }
            */

            // Todo: Check row tests

            var classname = parent.Name;
            var membername = test.Name;

            return _vbaTestSuite.GetTestClassMemberInfo(classname, membername);
        }

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();
            }

            DisposeUnmanagedResources();
            _disposed = true;
        }

        private void DisposeUnmanagedResources()
        {
            OfficeApplicationHelper = null;
        }

        private void DisposeManagedResources()
        {
            DisposeVbaTestSuite();
        }

        private void DisposeVbaTestSuite()
        {
            if (_vbaTestSuite == null)
                return;

            using (new BlockLogger())
            {
                try
                {
                    _vbaTestSuite.Dispose();
                    Logger.Log("_vbaTestSuite disposed");
                }
                catch (Exception exception)
                {
                    Logger.Log(exception);
                }
                finally
                {
                    _vbaTestSuite = null;
                }
            }   
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~TestSuiteManager()
        {
            Dispose(false);
        }
        #endregion

    }
}
