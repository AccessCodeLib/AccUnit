using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Configuration
{
    [ComVisible(true)]
    [Guid("1D30999D-85C3-4732-B2AC-E9E53F528241")]
    public interface IConfigurator
    {
        [ComVisible(false)]
        void Init(VBProject VBProject);

        [ComVisible(true)]
        void RemoveTestEnvironment(bool RemoveTestModules = false, bool ExportModulesBeforeRemoving = true, [MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);
        void InsertAccUnitLoaderFactoryModule(bool UseAccUnitTypeLib, bool RemoveIfExists = false, [MarshalAs(UnmanagedType.IDispatch)] object VBProject = null, object HostApplication = null);
        void RemoveAccUnitLoaderFactoryModule([MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);
        void ExportTestClasses(string ExportPath = null, [MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);
        void ImportTestClasses(string FileNameFilter = null, string ImportPath = null, [MarshalAs(UnmanagedType.IDispatch)] object VBProject = null);

        IUserSettings UserSettings { get; } 
    }

    [ComVisible(true)]
    [Guid("EE61931F-2E7C-4EA4-8C4B-B4E185336FEE")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.Configurator")]
    public class Configurator : IConfigurator, IDisposable
    {
        private VBProject _vbProject;

        public Configurator()
        {
        }

        public Configurator(VBProject vbproject)
        {
            _vbProject = vbproject;
        }

        public void InsertAccUnitLoaderFactoryModule(bool UseAccUnitTypeLib = false, bool removeIfExists = false, 
                    object vbProject = null, object HostApplication = null)
        {
            if (vbProject != null)
            {
                _vbProject = (VBProject)vbProject;
            }

            string hostName = "Microsoft Access";

            if (HostApplication != null)
            {
                OfficeApplicationHelper officeApplicationHelper = new OfficeApplicationHelper(HostApplication);
                hostName = officeApplicationHelper.Name;
            }

            var accUnitLoaderAddInCodeTemplates = new AccUnitLoaderAddInCodeTemplates(UseAccUnitTypeLib, hostName);

            if (removeIfExists)
            {
                try
                {
                    accUnitLoaderAddInCodeTemplates.RemoveFromVBProject(_vbProject);
                }
                catch { }
            }

            accUnitLoaderAddInCodeTemplates.EnsureModulesExistIn(_vbProject);
        }

        public void RemoveAccUnitLoaderFactoryModule(object vbProject = null)
        {
            if (vbProject != null)
                _vbProject = (VBProject)vbProject;

            var accUnitLoaderAddInCodeTemplates = new AccUnitLoaderAddInCodeTemplates(false);
            accUnitLoaderAddInCodeTemplates.RemoveFromVBProject(_vbProject);
        }

        public void Init(VBProject vbProject)
        {
            _vbProject = vbProject;

            //References.EnsureReferencesExistIn(_vbProject);
            //TestSuiteCodeTemplates.EnsureModulesExistIn(_vbProject);
        }

        public void RemoveTestEnvironment(bool removeTestModules = false, bool exportModulesBeforeRemoving = true, object vbProject = null)
        {
            if (vbProject != null)
                _vbProject = (VBProject)vbProject;

            if (removeTestModules)
            {
                OfficeApplicationHelper officeApplicationHelper = new VBProjectOnlyApplicatonHelper(_vbProject);
                using (var testClassManager = new TestClassManager(officeApplicationHelper))
                {
                    testClassManager.RemoveTestComponents(exportModulesBeforeRemoving);
                }
            }

            RemoveAccUnitLoaderFactoryModule();
            RemoveAccUnitTlbReference();
        }

        private void RemoveAccUnitTlbReference()
        {
            string refName;
            foreach (Reference reference in _vbProject.References)
            {
                try
                {
                    refName = reference.Name;
                }
                catch {
                    refName = "";
                }

                if (!string.IsNullOrEmpty(refName) && refName.Equals("AccUnit", StringComparison.CurrentCultureIgnoreCase))
                {
                    _vbProject.References.Remove(reference);
                    break;
                }
            }
        }

        public void ExportTestClasses(string exportPath = null, object vbProject = null)
        {
            if (vbProject != null)
                _vbProject = (VBProject)vbProject;

            OfficeApplicationHelper officeApplicationHelper = new VBProjectOnlyApplicatonHelper(_vbProject);
            using (var testClassManager = new TestClassManager(officeApplicationHelper))
            {
                testClassManager.ExportTestClasses(exportPath);
            }
        }

        public void ImportTestClasses(string FileNameFilter = null, string importPath = null, object VBProject = null)
        {
            OfficeApplicationHelper officeApplicationHelper = new VBProjectOnlyApplicatonHelper(_vbProject);
            using (var testClassManager = new TestClassManager(officeApplicationHelper))
            {
                testClassManager.ImportTestComponents(FileNameFilter, importPath, true);
            }
        }

        public IUserSettings UserSettings
        {
            get
            {
                return Configuration.UserSettings.Current;
            }
        }   

        /*
        public static void CheckAccUnitVBAReferences(VBProject vbProject)
        {
            throw new NotImplementedException("TestSuite-Factory ist noch nicht fertig");
            // var references = new AccUnitVBAReferences();
            // references.EnsureReferencesExistIn(vbProject);
        }
        */

        /*
        private TestSuiteCodeTemplates TestSuiteCodeTemplates { get; } = new TestSuiteCodeTemplates();

        private void DeleteFactoryCodeModule()
        {
            var factory = new TestClassFactoryManager(_vbProject, new TestClassReader(_vbProject));
            factory.DeleteFactoryCodeModule();
        }
        */

        #region IDisposable Support

        public delegate void DisposeEventHandler(object sender);
        public event DisposeEventHandler Disposed;

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                try
                {

                    if (disposing)
                    {
                        //
                    }

                    // unmanaged resources:
                    _vbProject = null;

                    _disposed = true;
                }
                catch (Exception ex) { Logger.Log(ex.Message); }

            }

            Disposed?.Invoke(this);

            // GC-Bereinigung wegen unmanaged res:
            // GC.Collect();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Configurator()
        {
            Dispose(false);
        }

        #endregion

    }
}
