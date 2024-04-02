using AccessCodeLib.AccUnit.Configuration;
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
    public class VBATestBuilder : IDisposable, IVBATestBuilder
    {
        private VBProject _vbProject;

        public VBATestBuilder(OfficeApplicationHelper applicationHelper)
        {
            OfficeApplicationHelper = applicationHelper;
            _vbProject = applicationHelper.CurrentVBProject;
            TestClassFactoryManager = new TestClassFactoryManager(_vbProject, new TestClassReader(_vbProject));
        }

        public event NullReferenceEventHandler OfficeApplicationReferenceRequired;

        private TestClassFactoryManager TestClassFactoryManager { get; set; }
        public bool TestToolsActivated { get; private set; }

        internal OfficeApplicationHelper OfficeApplicationHelper { get; private set; }

        private void CheckTestManagerInterface(object testToAdd, ITestClassMemberList memberFilter)
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

        public object CreateTest(object testToAdd, ITestClassMemberList memberFilter)
        {
            CheckTestManagerInterface(testToAdd, memberFilter);
            return testToAdd;
        }

        public object CreateTest(string className)
        {
            return CreateTest(className, null);
        }

        private object CreateTest(string className, ITestClassMemberList memberFilter)
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

        private void InitTestManager(ITestManagerBridge testToAdd, ITestClassMemberList memberFilter = null)
        {
            new TestManager(testToAdd, memberFilter) { ActiveVBProject = _vbProject, HostApplication = OfficeApplicationHelper.Application };
        }

        public object CreateObject(string className)
        {
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
            var modules = new CodeModuleContainer(_vbProject);
            var module = modules.TryGetCodeModule(className);
            if (module is null)
                return;

            if (!TestMessageBox.UsedInCodeModule(module)) return;

            TestMessageBox.CheckTestMessageBoxProcedures(_vbProject);
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
                if (OfficeApplicationHelper.Name.Equals("Microsoft Excel", StringComparison.CurrentCultureIgnoreCase))
                {
                    parameters[0] = GetFullRunMethodeNameForExcel(parameters[0].ToString());
                }
                return OfficeApplicationHelper.Run(parameters);
            }
            catch (Exception xcp)
            {
                throw new OfficeApplicationRunException(xcp, parameters);
            }
        }

        private string GetFullRunMethodeNameForExcel(string methodName)
        {
            var invokeHelper = new InvocationHelper(OfficeApplicationHelper.Application);
            var wb = invokeHelper.InvokePropertyGet("ActiveWorkbook");
            invokeHelper = new InvocationHelper(wb);
            var wbName = invokeHelper.InvokePropertyGet("Name");

            return string.Concat("'", wbName, "'!", methodName);
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
