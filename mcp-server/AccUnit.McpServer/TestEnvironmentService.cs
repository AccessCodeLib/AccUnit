using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Removes the AccUnit test environment from the open database, like the menu item of the VBE add-in.
    /// </summary>
    public sealed class TestEnvironmentService
    {
        private const string TestClassFactoryModuleName = "AccUnit_TestClassFactory";

        private readonly AccessSession _session;

        public TestEnvironmentService(AccessSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public RemoveTestEnvironmentResult RemoveTestEnvironment(bool removeTestModules, bool exportBeforeRemoving)
        {
            var result = _session.Execute(application =>
                RemoveOnStaThread(VBProjectLocator.GetCurrentVBProject(application), removeTestModules, exportBeforeRemoving));

            // The modules are removed with VBComponents.Remove (AccUnit core). Access keeps that only if the
            // VBA project is saved, as when closing Access interactively, so save and open the database again.
            _session.SaveAndReopen();
            result.Message = "Removed and saved; the database was closed with saving and opened again.";
            return result;
        }

        private static RemoveTestEnvironmentResult RemoveOnStaThread(VBProject vbProject, bool removeTestModules, bool exportBeforeRemoving)
        {
            var componentsBefore = GetComponentNames(vbProject);
            var exportFolder = removeTestModules && exportBeforeRemoving ? GetExportFolder(vbProject) : null;

            // Removes AccUnit_Factory and the AccUnit type library reference; with removeTestModules also the test classes.
            using (var configurator = new Configurator(vbProject))
            {
                configurator.RemoveTestEnvironment(removeTestModules, exportBeforeRemoving, vbProject);
            }

            // The MCP server adds AccUnit_TestClassFactory for test runs; Configurator removes it only together with the test classes.
            if (GetComponentNames(vbProject).Contains(TestClassFactoryModuleName, StringComparer.OrdinalIgnoreCase))
            {
                using (var factoryManager = new TestClassFactoryManager(vbProject, new TestClassReader(vbProject)))
                {
                    factoryManager.DeleteFactoryCodeModule();
                }
            }

            return new RemoveTestEnvironmentResult
            {
                RemovedModules = componentsBefore.Except(GetComponentNames(vbProject), StringComparer.OrdinalIgnoreCase).ToList(),
                ExportFolder = exportFolder
            };
        }

        private static string GetExportFolder(VBProject vbProject)
        {
            using (var testClassManager = new TestClassManager(new VBProjectOnlyApplicatonHelper(vbProject)))
            {
                return testClassManager.ExportDirectory;
            }
        }

        private static List<string> GetComponentNames(VBProject vbProject)
        {
            var names = new List<string>();
            foreach (VBComponent component in vbProject.VBComponents)
                names.Add(component.Name);
            return names;
        }
    }

    public class RemoveTestEnvironmentResult
    {
        public List<string> RemovedModules { get; set; } = new List<string>();

        /// <summary>Folder with the exported test classes (only if test classes were removed and exported).</summary>
        public string ExportFolder { get; set; }

        public string Message { get; set; }
    }
}
