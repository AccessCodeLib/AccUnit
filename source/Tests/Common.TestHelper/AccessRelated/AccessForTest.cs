using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace AccessCodeLib.Common.TestHelpers.AccessRelated
{
    public class AccessForTest
    {
        private const string _locale = ";LANGID=0x0409;CP=1252;COUNTRY=0";
        private class ProcessEqualityComparer : IEqualityComparer<Process>
        {
            public bool Equals(Process x, Process y)
            {
                return x.Id == y.Id;
            }

            public int GetHashCode(Process process)
            {
                return process.Id.GetHashCode();
            }
        }

        private bool _isDisposed;

        public AccessForTest()
        {
            StartAccessInstance();
            CreateAndOpenDatabase();
        }

        private void CreateAndOpenDatabase()
        {
            DBEngine = (DBEngine)Application.DBEngine;

            DatabaseFileFullPath = GetTempFileName();
            CreateDatabase(DatabaseFileFullPath);
            OpenDatabase(DatabaseFileFullPath);

            UpdateLockFileStatus();
        }

        private void OpenDatabase(string fullPath)
        {
            Application.OpenCurrentDatabase(fullPath);
            LockFileFullPath = GetLockFileFullPath(fullPath);
        }

        public string LockFileFullPath { get; set; }

        private string GetLockFileFullPath(string fullPath)
        {
            var fileInfo = new FileInfo(fullPath);

            var lockFilePath = Path.ChangeExtension(fileInfo.FullName, "ldb");
            if (File.Exists(lockFilePath))
            {
                return lockFilePath;
            }

            lockFilePath = Path.ChangeExtension(fileInfo.FullName, "laccdb");
            if (File.Exists(lockFilePath))
            {
                return lockFilePath;
            }
            throw new Exception("Unable to detect lock file.");
        }

        private void CreateDatabase(string fullPath)
        {
            //using (var wrappedDatabase = new ComWrapper<Database>(DBEngine.CreateDatabase(fullPath, _locale)))
            using (var wrappedDatabase = new ComWrapper<Database>(DBEngine.CreateDatabase(fullPath, _locale)))
            {
                var database = wrappedDatabase.ComReference;
                database.Close();
            }
        }

        private void UpdateLockFileStatus()
        {
            IsLockFilePresent = File.Exists(LockFileFullPath);
        }

        private DBEngine DBEngine { get; set; }

        private static string GetTempFileName()
        {
            return Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        }

        public bool HasExited
        {
            get { return Process.HasExited; }
        }

        private dynamic _application;

        public dynamic Application
        {
            get
            {
                EnsureNotDisposed();
                return _application;
            }
            private set { _application = value; }
        }

        private void EnsureNotDisposed()
        {
            if (_isDisposed)
                throw new ObjectDisposedException(GetType().FullName);
        }

        ~AccessForTest()
        {
            Dispose();
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;
            Debug.WriteLine(Process.Id + ": Dispose()");
            if (!Process.HasExited)
            {
                Process.Kill();
                Process.WaitForExit();
            }
            Process.Dispose();
            Process = null;
            Application = null;
            CleanupDatabaseFiles();
            _isDisposed = true;
        }

        private void CleanupDatabaseFiles()
        {
            if (File.Exists(DatabaseFileFullPath))
                File.Delete(DatabaseFileFullPath);
            if (File.Exists(LockFileFullPath))
                File.Delete(LockFileFullPath);
        }

        private void StartAccessInstance()
        {
            var accessProcessesBefore = GetRunningAccessInstances();

            Application = AccessFactory.CreateApplication();

            var accessProcessesAfter = GetRunningAccessInstances();
            var differenceProcesses = accessProcessesAfter.Except(accessProcessesBefore,
                                                                  new ProcessEqualityComparer());

            var differenceCount = differenceProcesses.ToList().Count;
            Debug.Assert(differenceCount > 0, "Could not find the process of the test Access application.");
            Debug.Assert(differenceCount == 1, "Invalid difference", "There is a difference of {0} Access processes.", differenceCount);

            Process = differenceProcesses.First();

            Trace.WriteLine(string.Format("********** This is thread \"{0}\". I just started access with process id #{1}",
                                          Thread.CurrentThread.Name, Process.Id));
        }

        private Process Process { get; set; }

        public Process GetProcess()
        {
            return Process.GetProcessById(Process.Id);
        }

        private static IEnumerable<Process> GetRunningAccessInstances()
        {
            var accessProcesses = Process.GetProcessesByName("MSACCESS");
            return accessProcesses;
        }

        public void Quit()
        {
            Quit(-1);
        }

        public void Quit(int timeout)
        {
            CloseDatabase();

            Marshal.ReleaseComObject(DBEngine);

            Application.Quit(AcQuitOption.acQuitSaveNone);
            Process.WaitForExit(timeout);
            Application = null;
        }

        public Microsoft.Office.Interop.Access.Dao.Database GetDatabase
        {
            get { return (Database)Application.CurrentDb(); }
        }

        public bool IsDatabaseOpen { get { return Application.CurrentDb() != null; } }

        public bool IsLockFilePresent { get; private set; }

        public string DatabaseFileFullPath { get; private set; }

        public VBE Vbe
        {
            get { return Application.VBE; }
        }

        public VBProject ActiveVbProject
        {
            get { return Vbe.ActiveVBProject; }
        }


        public void CloseDatabase()
        {
            Application.CloseCurrentDatabase();
            UpdateLockFileStatus();
        }

        public VBComponent AddClass(string className)
        {
            var newClass = ActiveVbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            newClass.Name = className;
            return newClass;
        }

        public void CompileAndSave()
        {
            Application.RunCommand(AcCommand.acCmdCompileAndSaveAllModules);
        }

        public void ShowAccess()
        {
            Application.Visible = true;
        }

        public TableBuilder AddTable(string tableName)
        {
            return new TableBuilder(() => Application.CurrentDb(), tableName);
        }

        public class TableBuilder
        {
            private readonly Func<Database> _getDatabase;
            private readonly string _tableName;
            private readonly IList<string> _fieldNames = new List<string>();

            public TableBuilder(Func<Database> getDatabase, string tableName)
            {
                _getDatabase = getDatabase;
                _tableName = tableName;
            }

            public void Create()
            {
                using (var databaseWrapper = new ComWrapper<Database>(_getDatabase()))
                {
                    var database = databaseWrapper.ComReference;
                    var newTable = database.CreateTableDef(_tableName);
                    if (_fieldNames.Count == 0)
                    {
                        _fieldNames.Add("dummy");
                    }
                    foreach (var fieldName in _fieldNames)
                    {
                        var theField = newTable.CreateField(fieldName, DataTypeEnum.dbLong);
                        newTable.Fields.Append(theField);
                    }
                    using (var tableDefsWrapper = new ComWrapper<TableDefs>(database.TableDefs))
                    {
                        tableDefsWrapper.ComReference.Append(newTable);
                        tableDefsWrapper.ComReference.Refresh();
                    }
                }
            }

            public TableBuilder AndField(string fieldName)
            {
                _fieldNames.Add(fieldName);
                return this;
            }
        }
    }
}