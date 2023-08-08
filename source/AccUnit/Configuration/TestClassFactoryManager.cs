using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;

namespace AccessCodeLib.AccUnit.Configuration
{
    public class TestClassFactoryManager : IDisposable
    {
        private const string FactoryCodeModuleName = "AccUnit_TestClassFactory";  // "_AccUnit_TestClassFactory" is not allowed in Excel
        private const string FactoryCodeModuleHeader = @"Option Compare Text
Option Explicit
Option Private Module
";

        private VBProject _vbProject;
        private CodeModule _factoryCodeModule;
        private readonly ITestClassReader _testClassReader;

        public TestClassFactoryManager(VBProject vbProject, ITestClassReader testClassReader)
        {
            _vbProject = vbProject;
            _testClassReader = testClassReader;
        }

        internal CodeModule FactoryModule
        {
            get
            {
                if (_factoryCodeModule is null)
                {
                    _factoryCodeModule = GetFactoryModule();
                }
                else
                {
                    try // maybe codemodule was deleted
                    {
                        // ReSharper disable UseIndexedProperty
                        _factoryCodeModule.get_Lines(1, 1);
                        // ReSharper restore UseIndexedProperty
                    }
                    catch
                    {
                        _factoryCodeModule = GetFactoryModule();
                    }
                }
                return _factoryCodeModule;
            }
        }

        private CodeModule GetFactoryModule()
        {
            using (var codemodulmanager = new CodeModuleContainer(_vbProject))
            {
                var codeModule = codemodulmanager.TryGetCodeModule(FactoryCodeModuleName) ??
                               codemodulmanager.Generator.Add(vbext_ComponentType.vbext_ct_StdModule, FactoryCodeModuleName, FactoryCodeModuleHeader, true);
                CloseCodePaneWindowSave(codeModule.CodePane);
                return codeModule;
            }
        }

        private static void CloseCodePaneWindowSave(_CodePane codePane)
        {
            try
            {
                codePane.Window.Close();
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch { /* Don't mind if there is an exception because the window is not open */ }
            // ReSharper restore EmptyGeneralCatchClause
        }

        private void AddFactoryMethod(string className)
        {
            FactoryModule.InsertLines(FactoryModule.CountOfLines + 1,
                                      GetFactoryMethodCode(className)
                                      );

        }

        internal static string GetFactoryMethodCode(string className)
        {
            return $@"Public Function {GetTestClassFactoryMethodName(className)}() As Object
   Set {GetTestClassFactoryMethodName(className)} = New {GetClassNameForNewStatement(className)}
End Function
";
        }

        private static string GetClassNameForNewStatement(string className)
        {
            return className.Contains(" ") || "_0123456789".Contains(className[0].ToString())
                ? $"[{className}]"
                : className;
        }

        public bool EnsureFactoryMethodExists(string className)
        {
            if (!FactoryMethodExists(className))
            {
                AddFactoryMethod(className);
                return true;
            }
            return false;
        }

        public static string GetTestClassFactoryMethodName(string className)
        {
            return $"AccUnitTestClassFactory_{className}".Replace(" ", "_");
        }

        /// @todo Remove catch all. Exceptions from FactoryModule may not be ignored.
        private bool FactoryMethodExists(string className)
        {
            try
            {
                var dummyLine = FactoryModule.ProcBodyLine[GetTestClassFactoryMethodName(className), vbext_ProcKind.vbext_pk_Proc];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void RefreshFactoryCodeModule()
        {
            // Brute force
            WipeFactoryCodeModule();
            RewriteFactoryCodeModule();
        }

        private void WipeFactoryCodeModule()
        {
            var vbc = FindFactoryVbComponent();
            if (vbc is null)
                return;

            var cm = vbc.CodeModule;
            cm.DeleteLines(1, cm.CountOfLines);
            InsertFactoryCodeModuleHeader();
        }

        private void RewriteFactoryCodeModule()
        {
            var testClasses = _testClassReader.GetTestClasses();
            foreach (var testClassName in testClasses.Select(tc => tc.Name).OrderBy(n => n))
            {
                AddFactoryMethod(testClassName);
            }
        }

        private void InsertFactoryCodeModuleHeader()
        {
            FactoryModule.InsertLines(1, FactoryCodeModuleHeader);
        }

        public void DeleteFactoryCodeModule()
        {
            var vbc = FindFactoryVbComponent();
            vbc?.Collection.Remove(vbc);
            _factoryCodeModule = null;
        }

        private VBComponent FindFactoryVbComponent()
        {
            try
            {
                var vbcCol = _vbProject.VBComponents;
                return vbcCol.Item(FactoryCodeModuleName);
            }
            catch (IndexOutOfRangeException)
            {
                return null;
            }
        }

        #region IDisposable Support
        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            try
            {
                DisposeUnmanagedResources();
                _disposed = true;
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }

        void DisposeUnmanagedResources()
        {
            _vbProject = null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~TestClassFactoryManager()
        {
            Dispose(false);
        }

        #endregion
    }
}