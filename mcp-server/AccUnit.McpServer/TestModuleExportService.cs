using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Exports single test classes and test related modules to the AccUnit test folder and removes them from the
    /// open database, with the functions of the AccUnit core. Other modules are never exported or removed.
    /// </summary>
    public sealed class TestModuleExportService
    {
        private readonly AccessSession _session;

        public TestModuleExportService(AccessSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public ExportTestClassesResult ExportTestClasses(IList<string> classNames, bool overwrite)
        {
            return _session.Execute(application =>
                ExportOnStaThread(VBProjectLocator.GetCurrentVBProject(application), classNames, overwrite));
        }

        public RemoveTestClassesResult RemoveTestClasses(IList<string> classNames, bool exportBeforeRemoving)
        {
            var result = _session.Execute(application =>
                RemoveOnStaThread(VBProjectLocator.GetCurrentVBProject(application), classNames, exportBeforeRemoving));

            if (result.RemovedModules.Count == 0)
            {
                result.Message = "Nothing was removed.";
                return result;
            }

            // Access keeps modules removed with VBComponents.Remove only if the VBA project is saved.
            _session.SaveAndReopen();
            result.Message = "Removed and saved; the database was closed with saving and opened again.";
            return result;
        }

        private static ExportTestClassesResult ExportOnStaThread(VBProject vbProject, IList<string> classNames, bool overwrite)
        {
            using (var testClassManager = new TestClassManager(new VBProjectOnlyApplicatonHelper(vbProject)))
            {
                var result = new ExportTestClassesResult { ExportDirectory = testClassManager.ExportDirectory };
                foreach (var module in SelectTestModules(vbProject, classNames, result.Rejected))
                {
                    var filePath = GetExportFilePath(result.ExportDirectory, module);
                    if (File.Exists(filePath) && !overwrite)
                    {
                        result.Rejected.Add(new RejectedTestModule
                        {
                            Name = module.Name,
                            Reason = "The file " + filePath + " exists already; export with overwrite = true (the file is kept as .bak)."
                        });
                        continue;
                    }

                    AddBackup(result.BackupFiles, SourceFileBackup.BackupIfExists(filePath));
                    result.ExportedFiles.Add(testClassManager.ExportTestClass(module.Name));
                }
                return result;
            }
        }

        private static RemoveTestClassesResult RemoveOnStaThread(VBProject vbProject, IList<string> classNames, bool exportBeforeRemoving)
        {
            var componentsBefore = GetComponentNames(vbProject);
            using (var testClassManager = new TestClassManager(new VBProjectOnlyApplicatonHelper(vbProject)))
            {
                var result = new RemoveTestClassesResult();
                foreach (var module in SelectTestModules(vbProject, classNames, result.Rejected))
                {
                    if (exportBeforeRemoving)
                    {
                        var filePath = GetExportFilePath(testClassManager.ExportDirectory, module);
                        AddBackup(result.BackupFiles, SourceFileBackup.BackupIfExists(filePath));
                        result.ExportedFiles.Add(filePath);
                    }
                    // Exports to the AccUnit test folder first if requested; also removes AccUnit_TestClassFactory,
                    // which accunit_run_tests adds again.
                    testClassManager.RemoveTestClass(module.Name, exportBeforeRemoving);
                }

                result.RemovedModules = componentsBefore.Except(GetComponentNames(vbProject), StringComparer.OrdinalIgnoreCase).ToList();
                return result;
            }
        }

        private static List<CodeModuleInfo> SelectTestModules(VBProject vbProject, IList<string> classNames, List<RejectedTestModule> rejected)
        {
            var testModules = new TestClassReader(vbProject).GetTestComponents().ToList();
            return TestModuleSelection.Select(classNames, testModules, GetComponentNames(vbProject), rejected);
        }

        private static string GetExportFilePath(string directory, CodeModuleInfo module)
        {
            return Path.Combine(directory, module.Name + TestModuleSelection.GetFileExtension(module.ComponentType));
        }

        private static void AddBackup(List<string> backupFiles, string backupFile)
        {
            if (backupFile != null)
                backupFiles.Add(backupFile);
        }

        private static List<string> GetComponentNames(VBProject vbProject)
        {
            var names = new List<string>();
            foreach (VBComponent component in vbProject.VBComponents)
                names.Add(component.Name);
            return names;
        }
    }

    public class ExportTestClassesResult
    {
        public string ExportDirectory { get; set; }
        public List<string> ExportedFiles { get; set; } = new List<string>();

        /// <summary>Copies (.bak) of files that the export wrote over.</summary>
        public List<string> BackupFiles { get; set; } = new List<string>();
        public List<RejectedTestModule> Rejected { get; set; } = new List<RejectedTestModule>();
    }

    public class RemoveTestClassesResult
    {
        /// <summary>All removed modules, including AccUnit_TestClassFactory (added again by accunit_run_tests).</summary>
        public List<string> RemovedModules { get; set; } = new List<string>();
        public List<string> ExportedFiles { get; set; } = new List<string>();
        public List<string> BackupFiles { get; set; } = new List<string>();
        public List<RejectedTestModule> Rejected { get; set; } = new List<RejectedTestModule>();
        public string Message { get; set; }
    }
}
