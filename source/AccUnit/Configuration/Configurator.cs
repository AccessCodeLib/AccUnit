using System;
using System.Runtime.InteropServices;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Configuration
{
    [ComVisible(true)]
    [Guid("1D30999D-85C3-4732-B2AC-E9E53F528241")]
    public interface IConfigurator
    {
        [ComVisible(false)]
        void Init(VBProject VBProject);
        [ComVisible(false)]
        void Remove(VBProject VBProject = null, bool RemoveTestModules = false, bool ExportModulesBeforeRemoving = true);

        [ComVisible(true)]
        void InsertAccUnitLoaderFactoryModule(VBProject VBProject, bool UseAccUnitTypeLib, bool removeIfExists = false);
        void RemoveAccUnitLoaderFactoryModule(VBProject VBProject);

    }

    [ComVisible(true)]
    [Guid("EE61931F-2E7C-4EA4-8C4B-B4E185336FEE")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.Configurator")]
    public class Configurator : IConfigurator, IDisposable
    {
        public Configurator()
        {
            //References = new AccUnitVBAReferences();
        }

        public Configurator(VBProject vbproject)
        {
            //References = new AccUnitVBAReferences();
            _vbProject = vbproject;
        }

        public void InsertAccUnitLoaderFactoryModule(VBProject vbProject = null, bool UseAccUnitTypeLib = false, bool removeIfExists = false)
        {
            if (vbProject != null)
            {
                _vbProject = vbProject;
            }
            var accUnitLoaderAddInCodeTemplates = new AccUnitLoaderAddInCodeTemplates(UseAccUnitTypeLib);

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

        public void RemoveAccUnitLoaderFactoryModule(VBProject vbProject = null)
        {
            if (vbProject != null)
                _vbProject = vbProject;

            var accUnitLoaderAddInCodeTemplates = new AccUnitLoaderAddInCodeTemplates(false);
            accUnitLoaderAddInCodeTemplates.RemoveFromVBProject(_vbProject);
        }

        public void Init(VBProject vbProject)
        {
            _vbProject = vbProject;
            
            //References.EnsureReferencesExistIn(_vbProject);
            //TestSuiteCodeTemplates.EnsureModulesExistIn(_vbProject);
        }

        public void Remove(VBProject vbProject = null, bool removeTestModules = false, bool exportModulesBeforeRemoving = true)
        {
            throw new NotImplementedException("TestSuite-Factory ist noch nicht fertig");
            
            /*
            if (vbProject != null)
                _vbProject = vbProject;
            
           
            
            if (removeTestModules)
            {
                OfficeApplicationHelper officeApplicationHelper = new VbeOnlyApplicatonHelper(_vbProject.VBE);
                using (var testClassManager = new TestClassManager(officeApplicationHelper))
                {
                    testClassManager.RemoveTestComponents(exportModulesBeforeRemoving);
                }
            }

            DeleteFactoryCodeModule();
            TestSuiteCodeTemplates.RemoveFromVBProject(_vbProject);
            References.RemoveReferencesFrom(_vbProject);
            */
        }


        public AccUnitVBAReferences References { get; private set; }

        public static void CheckAccUnitVBAReferences(VBProject vbProject)
        {
            throw new NotImplementedException("TestSuite-Factory ist noch nicht fertig");
            // var references = new AccUnitVBAReferences();
            // references.EnsureReferencesExistIn(vbProject);
        }

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
        private VBProject _vbProject;

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

            if (Disposed != null)
            {
                Disposed(this);
            }

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
