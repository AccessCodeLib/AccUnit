using System;
using System.Collections.Generic;
using AccessCodeLib.AccUnit.Configuration;
using Microsoft.Vbe.Interop;
using Access = Microsoft.Office.Interop.Access;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Prepares the VBA project of the open database so that the tests can be run from outside Access.
    /// Must be used on the STA thread of the session.
    /// </summary>
    internal sealed class TestEnvironment
    {
        private const Access.AcCommand CompileAndSaveAllModules = (Access.AcCommand)126; // acCmdCompileAndSaveAllModules
        private const string FactoryModuleName = "AccUnit_Factory";

        private readonly Access.Application _application;
        private readonly VBProject _vbProject;

        public TestEnvironment(Access.Application application, VBProject vbProject)
        {
            _application = application;
            _vbProject = vbProject;
        }

        public List<string> Changes { get; } = new List<string>();

        /// <summary>Compiles and saves all modules. A compile error does not show a dialog.</summary>
        public bool Compile()
        {
            _application.RunCommand(CompileAndSaveAllModules);
            return _application.IsCompiled;
        }

        /// <summary>
        /// Test classes use Assert/Iz and the test suite from AccUnit_Factory, as in the Immediate window.
        /// </summary>
        public void EnsureFactoryModule()
        {
            if (ComponentExists(FactoryModuleName))
                return;

            using (var configurator = new Configurator())
            {
                configurator.InsertAccUnitLoaderFactoryModule(false, false, _vbProject, _application);
            }
            Changes.Add("Inserted module " + FactoryModuleName + ".");
        }

        /// <summary>
        /// AccUnit creates the test class instances through factory methods in AccUnit_TestClassFactory and adds
        /// missing methods while the tests are added. Changing the VBA project resets it after a call from outside
        /// Access, which destroys the new instances. So the methods are created before the run.
        /// </summary>
        public void EnsureTestClassFactoryMethods(IEnumerable<string> classNames)
        {
            var added = false;
            using (var factoryManager = new TestClassFactoryManager(_vbProject, new TestClassReader(_vbProject)))
            {
                foreach (var className in classNames)
                    added |= factoryManager.EnsureFactoryMethodExists(className);
            }
            if (added)
                Changes.Add("Added factory methods to module AccUnit_TestClassFactory.");
        }

        private bool ComponentExists(string name)
        {
            foreach (VBComponent component in _vbProject.VBComponents)
            {
                if (string.Equals(component.Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}
