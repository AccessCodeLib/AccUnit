using AccessCodeLib.AccUnit.Configuration;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.Templates;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit
{

    public class TestClassManager : IDisposable
    {

        private struct OfficeObjectInfo
        {
            public string Name;
            public AccessApplicationHelper.AcObjectType ObjectType;
            public string FileExtension;
        }

        public event EventHandler RepairActiveVBProjectCOMException;
        public event EventHandler ScanningForTestModules;

        private const string DefaultExportDir = @"%APPFOLDER%\Tests\%APPNAME%";

        public TestClassManager()
        {
        }

        public TestClassManager(OfficeApplicationHelper applicationHelper)
        {
            ApplicationHelper = applicationHelper;
        }

        public VBProject ActiveVBProject { get { return ApplicationHelper.CurrentVBProject; } }
        public OfficeApplicationHelper ApplicationHelper { get; set; }

        #region CodeModule Support

        private CodeModuleContainer _codeModuleManager;
        private CodeModuleContainer CodeModuleManager
        {
            get { return _codeModuleManager ?? (_codeModuleManager = new CodeModuleContainer(ActiveVBProject)); }
        }

        public string ExportTestClass(string name)
        {
            return ExportTestClass(name, ExportDirectory);
        }

        private string ExportTestClass(string name, string exportdirectory)
        {
            return CodeModuleManager.Export(name, exportdirectory);
        }

        public string ExportTestClass(VBComponent component)
        {
            return CodeModuleManager.Export(component, ExportDirectory);
        }

        public void RemoveTestClass(string name, bool export = true)
        {
            RemoveVbComponent(name, export, export ? ExportDirectory : null);
            DeleteFactoryCodeModule();
        }

        public void RemoveTestComponents(bool export = true, bool removeTestEnvironment = false)
        {
            var components = new TestClassReader(ActiveVBProject).GetTestComponents();
            RemoveTestComponents(components, export);

            if (removeTestEnvironment)
            {
                using (var configurator = new Configurator(ActiveVBProject))
                {
                    configurator.RemoveTestEnvironment();
                }
            }
        }

        public void RemoveTestComponents(IEnumerable<CodeModuleInfo> list, bool export = true)
        {
            var exportDirectory = export ? ExportDirectory : null;
            foreach (var c in list)
            {
                if (c.ComponentType == vbext_ComponentType.vbext_ct_Document)
                    RemoveOfficeDocument(c.Name, export, exportDirectory);
                else
                    RemoveVbComponent(c.Name, export, exportDirectory);
            }
            DeleteFactoryCodeModule();
        }

        private void RemoveVbComponent(string name, bool export, string exportdirectory)
        {
            if (export)
            {
                CodeModuleManager.ExportAndRemove(name, exportdirectory);
            }
            else
            {
                CodeModuleManager.Remove(name);
            }
        }

        private void RemoveOfficeDocument(string name, bool export, string exportdirectory)
        {
            if (!(ApplicationHelper is AccessApplicationHelper accessApplication))
                return;

            var objectInfo = GetObjectInfo(name);

            if (export)
            {
                var fileName = $"{exportdirectory.TrimEnd(' ', '\\')}\\{name}{objectInfo.FileExtension}";
                accessApplication.SaveAsText(objectInfo.ObjectType, objectInfo.Name, fileName);
            }

            accessApplication.DoCmd.DeleteObject(objectInfo.ObjectType, objectInfo.Name);
        }

        private static OfficeObjectInfo GetObjectInfo(string name)
        {
            using (new BlockLogger())
            {
                var objectInfo = new OfficeObjectInfo
                {
                    Name = name,
                    FileExtension = "aco",
                    ObjectType = AccessApplicationHelper.AcObjectType.AcDefault
                };
                if (objectInfo.Name.IndexOf("Form_", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    objectInfo.Name = objectInfo.Name.Substring(5);
                    objectInfo.FileExtension = ".acf";
                    objectInfo.ObjectType = AccessApplicationHelper.AcObjectType.AcForm;
                }
                else if (objectInfo.Name.IndexOf("Report_", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    objectInfo.Name = objectInfo.Name.Substring(7);
                    objectInfo.FileExtension = ".acr";
                    objectInfo.ObjectType = AccessApplicationHelper.AcObjectType.AcReport;
                }
                return objectInfo;
            }
        }

        public void ExportTestClasses(string exportDirectory = null)
        {
            if (exportDirectory == null)
            {
                exportDirectory = ExportDirectory;
            }

            var classNames = new TestClassReader(ActiveVBProject).GetTestClasses();
            foreach (var testClassInfo in classNames)
            {
                ExportTestClass(testClassInfo.Name, exportDirectory);
            }
        }

        public void DeleteFactoryCodeModule()
        {
            var factory = new TestClassFactoryManager(ActiveVBProject, new TestClassReader(ActiveVBProject));
            factory.DeleteFactoryCodeModule();
        }

        private string _exportDir;
        public string ExportDirectory
        {
            get
            {
                if (string.IsNullOrEmpty(_exportDir))
                {
                    _exportDir = TestSuiteUserSettings.Current.ImportExportFolder;
                    if (string.IsNullOrEmpty(_exportDir))
                    {
                        _exportDir = DefaultExportDir;
                        TestSuiteUserSettings.Current.ImportExportFolder = _exportDir;
                    }
                }
                return GetCheckedFolder(_exportDir);
            }
            set
            {
                _exportDir = GetCheckedFolder(value);
            }
        }

        private string _importDir;
        public string ImportDirectory
        {
            get
            {
                if (string.IsNullOrEmpty(_importDir))
                {
                    _importDir = !string.IsNullOrEmpty(_exportDir) ? _exportDir : ExportDirectory;
                }
                return GetCheckedFolder(_importDir);
            }
            set
            {
                _importDir = GetCheckedFolder(value);
            }
        }

        internal string ApplicationPath
        {
            get
            {
                return Directory.GetParent(ActiveVBProject.FileName).ToString();
            }
        }

        private string ApplicationName
        {
            get
            {
                var fileInfo = new FileInfo(ActiveVBProject.FileName);
                return fileInfo.Name;
            }
        }

        private static readonly Regex AppFolderReplaceRegex = new Regex("%APPFOLDER%", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex AppNameReplaceRegex = new Regex("%APPNAME%", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);

        internal string GetCheckedFolder(string path)
        {
            path = AppFolderReplaceRegex.Replace(path, ApplicationPath);
            path = AppNameReplaceRegex.Replace(path, ApplicationName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        public void ImportTestComponents(string fileNameFilter = null, string importPath = null, bool overwriteExistingComponent = false)
        {
            var importDirectory = importPath ?? ImportDirectory;
            if (fileNameFilter == null)
            {
                fileNameFilter = "*";
            }
            var importComponents = GetTestModulesFromDirectory(importDirectory, fileNameFilter);

            ImportTestComponents(importComponents, overwriteExistingComponent);
        }

        public void ImportTestComponents(IEnumerable<CodeModuleInfo> list, bool overwriteExistingComponent = false)
        {
            foreach (var m in list)
            {
                if (overwriteExistingComponent)
                {
                    RemoveComponentIfExists(m);
                }
                ImportTestComponent(new FileInfo(m.FileName), m.ComponentType);
            }
        }

        private void RemoveComponentIfExists(CodeModuleInfo codeModuleInfo)
        {
            if (codeModuleInfo.ComponentType == vbext_ComponentType.vbext_ct_Document)
            {
                try
                {
                    RemoveOfficeDocument(codeModuleInfo.Name, false, null);
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            else
            {
                CodeModuleManager.Remove(codeModuleInfo.Name);
            }
        }


        private void ImportTestComponent(FileInfo importFile, vbext_ComponentType type)
        {
            if (type == vbext_ComponentType.vbext_ct_Document)
                ImportOfficeObject(importFile);
            else
            {
                CodeModuleManager.Generator.Add(importFile);
            }
        }

        private void ImportOfficeObject(FileSystemInfo importFile)
        {
            if (!(ApplicationHelper is AccessApplicationHelper accessApplication))
                return;

            var name = importFile.Name;
            name = name.Substring(0, name.Length - importFile.Extension.Length);
            var objectInfo = GetObjectInfo(name);
            accessApplication.LoadFromText(objectInfo.ObjectType, objectInfo.Name, importFile.FullName);
        }

        private IEnumerable<FileInfo> GetTestFilesFromDirectory(string path = null, string fileNameSeachPattern = "*")
        {
            if (path == null)
                path = ImportDirectory;
            return TestClassReader.GetTestFilesFromDirectory(path, fileNameSeachPattern);
        }

        public IEnumerable<CodeModuleInfo> GetTestModulesFromDirectory(string path = null, string fileNameSeachPattern = "*")
        {
            if (path == null)
                path = ImportDirectory;

            return from FileInfo file in GetTestFilesFromDirectory(path, fileNameSeachPattern)
                   select CreateCodeModuleInfo(file);
        }

        private static CodeModuleInfo CreateCodeModuleInfo(FileSystemInfo file)
        {
            return new CodeModuleInfo(file);
        }

        #endregion

        public IEnumerable<CodeModuleInfo> GetTestModulesFromVBProject()
        {
            using (new BlockLogger())
            {
                if (ActiveVBProject == null)
                {
                    Logger.Log("ActiveVBProject is null, raise RepairActiveVBProjectCOMException");
                    RaiseRepairActiveVBProjectCOMException();
                }

                TestClassReader testClassReader;

                try
                {
                    // TODO: How can this fail?
                    testClassReader = new TestClassReader(ActiveVBProject);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);

                    RaiseRepairActiveVBProjectCOMException();

                    testClassReader = new TestClassReader(ActiveVBProject);
                }

                return testClassReader.GetTestComponents();
            }
        }

        private void RaiseScanningForTestModules()
        {
            ScanningForTestModules?.Invoke(this, EventArgs.Empty);
        }

        private void RaiseRepairActiveVBProjectCOMException()
        {
            RepairActiveVBProjectCOMException?.Invoke(this, EventArgs.Empty);
        }

        public TestClassList GetTestClassListFromVBProject(bool readMemberInfo)
        {
            using (new BlockLogger())
            {
                if (ActiveVBProject == null)
                {
                    Logger.Log("ActiveVBProject is null, raise RepairActiveVBProjectCOMException");
                    RaiseRepairActiveVBProjectCOMException();
                }

                TestClassList list;
                try // issue #37
                {
                    list = GetTestClassListFromVBProject();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);

                    RaiseRepairActiveVBProjectCOMException();

                    list = GetTestClassListFromVBProject();
                }

                if (readMemberInfo)
                {
                    ReadTestClassMemberInfo(list);
                }

                return list;
            }
        }

        public TestClassList GetTestClassListFromVBProject(IEnumerable<TestItemTag> filterTags)
        {
            var list = GetTestClassListFromVBProject(true);
            if (filterTags != null && list != null)
            {
                list = new TestClassList(list.Where(c => c.IsMatch(filterTags)));
            }
            return list;
        }

        private TestClassList GetTestClassListFromVBProject()
        {
            using (new BlockLogger())
            {
                RaiseScanningForTestModules();
                var testClassReader = new TestClassReader(ActiveVBProject);
                return testClassReader.GetTestClasses();
            }
        }

        private void ReadTestClassMemberInfo(IEnumerable<TestClassInfo> list)
        {
            foreach (var c in list)
            {
                if (c.Members != null) continue;
                var reader = new TestClassReader(ActiveVBProject);
                c.InitMembers(reader.GetTestClassMemberList(c.Name));
            }
        }

        public TestClassInfo FindFirstMissingTestClassInVBProject(IEnumerable<TestClassInfo> list)
        {
            var completeList = GetTestClassListFromVBProject(false);
            return list.FirstOrDefault(testClassInfo => !completeList.Exists(x => x.Name == testClassInfo.Name));
        }

        public CodeModule InsertTestTemplate(CodeTemplate template, string templateName)
        {
            return template.AddToVBProject(ActiveVBProject, templateName);
        }

        #region IDisposable Support

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    if (_codeModuleManager != null)
                    {
                        _codeModuleManager.Dispose();
                        _codeModuleManager = null;
                    }

                    ApplicationHelper = null;
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }

            _disposed = true;

            // GC-Bereinigung wegen unmanaged res:
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            //GC.Collect();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~TestClassManager()
        {
            Dispose(false);
        }

        #endregion

    }

}
