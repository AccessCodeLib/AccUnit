using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Integration;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.AccUnit
{
    public class VBATestBuilder : IDisposable
    {
        private VBProject _vbProject;

        public event NullReferenceEventHandler OfficeApplicationReferenceRequired;

        private TestClassFactoryManager TestClassFactoryManager { get; set; }
        public bool TestToolsActivated { get; private set; }

        internal OfficeApplicationHelper OfficeApplicationHelper { get; private set; }

        public object HostApplication
        {
            get { return OfficeApplicationHelper?.Application; }
            set
            {
                OfficeApplicationHelper = ComTools.GetTypeForComObject(value, "Access.Application") != null
                                                ? new AccessApplicationHelper(value) : new OfficeApplicationHelper(value);

                _vbProject = OfficeApplicationHelper.CurrentVBProject;
                TestClassFactoryManager = new TestClassFactoryManager(_vbProject, new TestClassReader(_vbProject));
            }
        }

        public VBProject ActiveVBProject
        {
            get
            {
                if (_vbProject is null && HostApplication != null)
                {
                    _vbProject = OfficeApplicationHelper.CurrentVBProject;
                }
                return _vbProject;
            }
            set
            {
                _vbProject = value;
                TestClassFactoryManager = new TestClassFactoryManager(_vbProject, new TestClassReader(_vbProject));
            }
        }

        private void CheckTestManagerInterface(object testToAdd, TestClassMemberList memberFilter)
        {
            if (testToAdd is ITestManagerBridge bridge)
            {
                InitTestManager(bridge, memberFilter);
            }
        }

        public IEnumerable<object> CreateTests(IEnumerable<TestClassInfo> testClasses)
        {
            return testClasses.Select(testClass => CreateTest(testClass.Name, testClass.Members));
        }

        public object CreateTest(object testToAdd, TestClassMemberList memberFilter)
        {
            CheckTestManagerInterface(testToAdd, memberFilter);
            return testToAdd;
        }

        public object CreateTest(string className)
        {
            return CreateTest(className, null);
        }

        private object CreateTest(string className, TestClassMemberList memberFilter)
        {
            var testToAdd = CreateObject(className);
            CheckTestManagerInterface(testToAdd, memberFilter);
            return testToAdd;
        }

        public IEnumerable<object> CreateTestsFromVBProject()
        {
            var testClassReader = new TestClassReader(_vbProject);
            var testClasses = testClassReader.GetTestClasses();
            return CreateTests(testClasses);
        }

        private void InitTestManager(ITestManagerBridge testToAdd, TestClassMemberList memberFilter = null)
        {
            new TestManager(testToAdd, memberFilter) { ActiveVBProject = ActiveVBProject, HostApplication = HostApplication };
        }

        public object CreateObject(string className)
        {
            EnsureOfficeApplicationExists();
            CheckTools(className);
            if (TestClassFactoryManager.EnsureFactoryMethodExists(className))
            {
                AccessApplicationHelper accessAppHelper = OfficeApplicationHelper as AccessApplicationHelper;
                accessAppHelper?.RunCommand(AccessApplicationHelper.AcCommand.AcCmdCompileAndSaveAllModules);
            }

            var factoryMethodName = TestClassFactoryManager.GetTestClassFactoryMethodName(className);

            return RunMethodInOfficeApplication(factoryMethodName);
        }

        public void DeleteFactoryCodeModule()
        {
            TestClassFactoryManager.DeleteFactoryCodeModule();
        }

        public void RefreshFactoryCodeModule()
        {
            TestClassFactoryManager.RefreshFactoryCodeModule();
        }

        private void CheckTools(string className)
        {
            var modules = new CodeModuleContainer(ActiveVBProject);
            var module = modules.TryGetCodeModule(className);
            if (module is null)
                return;

            if (!TestMessageBox.UsedInCodeModule(module)) return;

            TestMessageBox.CheckTestMessageBoxProcedures(ActiveVBProject);
            TestToolsActivated = true;
        }

        private object RunMethodInOfficeApplication(object parameter)
        {
            return RunMethodInOfficeApplication(new[] { parameter });
        }

        private object RunMethodInOfficeApplication(object[] parameters)
        {
            try
            {
                return OfficeApplicationHelper.Run(parameters);
            }
            catch (Exception xcp)
            {
                throw new OfficeApplicationRunException(xcp, parameters);
            }
        }

        private void EnsureOfficeApplicationExists()
        {
            if (HostApplication is null)
            {
                HostApplication = GetOfficeApplication();
            }
        }

        private object GetOfficeApplication()
        {
            var app = QueryOfficeApplication() ?? throw new NullReferenceException("Office application reference");
            return app;
        }

        private object QueryOfficeApplication()
        {
            object app = null;
            OfficeApplicationReferenceRequired?.Invoke(ref app);
            return app;
        }

        #region IDisposable Support
        bool _disposed;
        protected void Dispose(bool disposing)
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
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }

        void DisposeManagedResources()
        {
            TestClassFactoryManager.Dispose();
            //ConstantsReader.Clear();
        }

        void DisposeUnmanagedResources()
        {
            _vbProject = null;
            OfficeApplicationHelper = null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VBATestBuilder()
        {
            Dispose(false);
        }

        #endregion
    }
}
