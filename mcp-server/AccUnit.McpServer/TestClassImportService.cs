using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using Access = Microsoft.Office.Interop.Access;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Imports test classes and test related modules from source files into the open database with the import of the
    /// AccUnit core (as the VBE add-in does). Other modules are never written.
    /// </summary>
    public sealed class TestClassImportService
    {
        private static readonly string[] ImportableExtensions = { ".cls", ".bas" };
        private static readonly char[] PathCharacters = { '\\', '/', ':' };
        private const string ParentDirectory = "..";

        private readonly AccessSession _session;

        public TestClassImportService(AccessSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        /// <summary>VBComponents.Import reads the system ANSI code page (Encoding.Default on .NET Framework).</summary>
        private static Encoding AnsiEncoding
        {
            get { return Encoding.Default; }
        }

        public TestClassImportResult ImportTestClasses(string fileNameFilter, string importPath)
        {
            var filter = NormalizeFileNameFilter(fileNameFilter);
            var directory = importPath == null ? null : ValidateImportPath(importPath);

            var result = _session.Execute(application =>
            {
                var importResult = ImportOnStaThread(VBProjectLocator.GetCurrentVBProject(application), filter, directory);
                SaveModules(application, importResult.ImportedModules);
                return importResult;
            });

            if (result.ImportedModules.Count == 0)
            {
                result.Message = "Nothing was imported.";
                return result;
            }

            // A module replaced with VBComponents.Remove is kept removed only if the VBA project is saved, so save and open again.
            _session.SaveAndReopen();
            result.Message = "Imported and saved; the database was closed with saving and opened again.";
            return result;
        }

        /// <summary>File name without extension, wildcards * and ? allowed, no path.</summary>
        public static string NormalizeFileNameFilter(string fileNameFilter)
        {
            if (string.IsNullOrWhiteSpace(fileNameFilter))
                throw new DatabaseSessionException("The file name filter is empty. Use the module name, wildcards * and ? are allowed.");

            var filter = fileNameFilter.Trim();
            if (filter.IndexOfAny(PathCharacters) >= 0 || filter.Contains(ParentDirectory))
                throw new DatabaseSessionException("The file name filter must not contain a path; use importPath for the folder.");

            var extension = Path.GetExtension(filter);
            return ImportableExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)
                ? filter.Substring(0, filter.Length - extension.Length)
                : filter;
        }

        public static string ValidateImportPath(string importPath)
        {
            string fullPath;
            try
            {
                fullPath = Path.IsPathRooted(importPath) ? Path.GetFullPath(importPath) : null;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                throw new DatabaseSessionException("The import path is not a valid folder path.", ex);
            }

            if (fullPath == null)
                throw new DatabaseSessionException("The import path must be an absolute path.");
            if (!Directory.Exists(fullPath))
                throw new DatabaseSessionException("The import folder does not exist: " + fullPath);
            return fullPath;
        }

        private static TestClassImportResult ImportOnStaThread(VBProject vbProject, string fileNameFilter, string importPath)
        {
            using (var testClassManager = new TestClassManager(new VBProjectOnlyApplicatonHelper(vbProject)))
            {
                var result = new TestClassImportResult { ImportDirectory = importPath ?? testClassManager.ImportDirectory };
                var files = GetSourceFiles(result.ImportDirectory, fileNameFilter);
                if (files.Count == 0)
                {
                    result.RejectedFiles.Add(new RejectedTestModuleFile
                    {
                        FileName = fileNameFilter + ".cls/.bas",
                        Reason = "No matching file in " + result.ImportDirectory + "."
                    });
                    return result;
                }

                var accepted = SelectImportableFiles(vbProject, files, result.RejectedFiles);
                if (accepted.Count > 0)
                    Import(testClassManager, accepted, result);
                return result;
            }
        }

        private static List<string> GetSourceFiles(string directory, string fileNameFilter)
        {
            return TestClassReader.GetTestFilesFromDirectory(directory, fileNameFilter)
                .Where(file => ImportableExtensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase))
                .Select(file => file.FullName)
                .ToList();
        }

        private static List<TestModuleSourceFile> SelectImportableFiles(VBProject vbProject, IEnumerable<string> files,
                                                                        List<RejectedTestModuleFile> rejectedFiles)
        {
            var componentNames = GetComponentNames(vbProject);
            var testModuleNames = new TestClassReader(vbProject).GetTestComponents().Select(module => module.Name).ToList();

            var accepted = new List<TestModuleSourceFile>();
            foreach (var filePath in files)
            {
                var file = TestModuleSourceFile.Read(filePath, AnsiEncoding);
                var reason = file.GetRejectReason(AnsiEncoding);
                if (reason == null && componentNames.Contains(file.ModuleName, StringComparer.OrdinalIgnoreCase) &&
                    !testModuleNames.Contains(file.ModuleName, StringComparer.OrdinalIgnoreCase))
                {
                    reason = "The database contains a module " + file.ModuleName + " that is not a test module; it is not replaced.";
                }

                if (reason == null)
                    accepted.Add(file);
                else
                    rejectedFiles.Add(new RejectedTestModuleFile { FileName = Path.GetFileName(filePath), Reason = reason });
            }
            return accepted;
        }

        private static void Import(TestClassManager testClassManager, List<TestModuleSourceFile> files, TestClassImportResult result)
        {
            var ansiDirectory = Path.Combine(Path.GetTempPath(), "AccUnitMcpServer", "Import-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(ansiDirectory);
            try
            {
                foreach (var file in files)
                    file.WriteAnsiCopy(ansiDirectory, AnsiEncoding);

                // Same import as the VBE add-in: an existing module with the same name is replaced.
                testClassManager.ImportTestComponents(
                    files.Select(file => new CodeModuleInfo(new FileInfo(Path.Combine(ansiDirectory, Path.GetFileName(file.FilePath))))),
                    true);
                result.ImportedModules.AddRange(files.Select(file => file.ModuleName));
            }
            finally
            {
                Directory.Delete(ansiDirectory, true);
            }
        }

        /// <summary>
        /// Saves each imported module without compiling (test classes compile only with AccUnit_Factory).
        /// Without it, quitting Access with acQuitSaveAll shows a modal "Save As" dialog for every imported module.
        /// </summary>
        private static void SaveModules(Access.Application application, IEnumerable<string> moduleNames)
        {
            var doCmd = application.DoCmd;
            try
            {
                foreach (var moduleName in moduleNames)
                    doCmd.Save(Access.AcObjectType.acModule, moduleName);
            }
            finally
            {
                Marshal.ReleaseComObject(doCmd);
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

    public class TestClassImportResult
    {
        public string ImportDirectory { get; set; }
        public List<string> ImportedModules { get; set; } = new List<string>();
        public List<RejectedTestModuleFile> RejectedFiles { get; set; } = new List<RejectedTestModuleFile>();
        public string Message { get; set; }
    }

    public class RejectedTestModuleFile
    {
        public string FileName { get; set; }
        public string Reason { get; set; }
    }
}
