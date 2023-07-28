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
    public class AccessTestHelper : IDisposable
    {
        private bool _isDisposed;
        private readonly string _locale;
        private readonly string _path;
        private const int MaxAttempt = 500;

        public AccessTestHelper(string locale = ";LANGID=0x0409;CP=1252;COUNTRY=0", string filePath = "")
        {
            DatabaseDeleteTimeout = 100;
            EnsureAccessKilledWaitInterval = 20;
            _locale = locale;
            _path = filePath;
            StartAccessTestInstance();
            CreateAndOpenDatabase();
        }

        public AccessTestHelper(string filePath) : this(";LANGID=0x0409;CP=1252;COUNTRY=0", filePath)
        { }

        public int EnsureAccessKilledWaitInterval { get; set; }

        ~AccessTestHelper()
        {
            Dispose();
        }

        public int DatabaseDeleteTimeout { get; set; }

        public void Dispose()
        {
            Cleanup();
            GC.SuppressFinalize(this);
            GC.Collect();

            //Thread.Sleep(EnsureAccessKilledWaitInterval);
            //EnsureAccessInstanceIsNotRunning();
        }

        private void EnsureAccessInstanceIsNotRunning()
        {
            if (IsTestInstanceRunning)
                throw new Exception("Could not shutdown test instance of Microsoft Access.");
        }

        public string Locale
        {
            get { return _locale; }
        }

        public VBComponent AddClassModule(string name)
        {
            var classModule = ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            classModule.Name = name;
            return classModule;
        }

        public dynamic Application { get; private set; }

        public Database CurrentDb { get { return (Database)Application.CurrentDb(); } }
        public ADODB.Connection CurrentConnection { get { return (ADODB.Connection)Application.CurrentProject.Connection; } }
        public CurrentProject CurrentProject { get { return (CurrentProject)Application.CurrentProject; } }

        public VBE VBE
        {
            get { return Application.VBE; }
        }

        public VBProject ActiveVBProject
        {
            get { return Application.VBE.ActiveVBProject; }
        }

        private void StartAccessTestInstance()
        {
            var accessProcessesBefore = GetRunningAccessInstancesProcessId();

            Application = AccessFactory.CreateApplication();
            //DBEngine = (DBEngine)Application.DBEngine; // nicht verfügbar?

            var accessProcessesAfter = GetRunningAccessInstancesProcessId();
            var differenceProcesses = GetDifference(accessProcessesBefore, accessProcessesAfter);

            Debug.Assert(differenceProcesses.Count > 0, "Could not find the process of the test Access application.");
            Debug.Assert(differenceProcesses.Count == 1, "Invalid difference", "There is a difference of {0} Access processes.", differenceProcesses.Count);

            ProcessId = differenceProcesses[0];

            Trace.WriteLine(string.Format("********** This is thread \"{0}\". I just started access with process id #{1}",
                                          Thread.CurrentThread.Name, ProcessId));
        }

        private void CreateAndOpenDatabase()
        {
            TempDbFileName = CreateTestDb();
            Trace.WriteLine(string.Format("Created testdatabase file in {0}", TempDbFileName));
            //accessApplication.OpenCurrentDatabase(dbFileName);
            //const string dbFileName = @"C:\Users\pro\AppData\Local\Temp\tmp1400.tmp";
            //accessApplication.DBEngine.OpenDatabase(dbFileName);
            Application.OpenCurrentDatabase(TempDbFileName);
        }

        private string TempDbFileName { get; set; }

        private _DBEngine _dbEngine;
        public _DBEngine DBEngine
        {
            get
            {
                if (_dbEngine is null)
                {
                    _dbEngine = (DBEngine)Application.DBEngine;
                }
                return _dbEngine;
            }
            private set
            {
                _dbEngine = value;
            }
        }

        public int ProcessId { get; private set; }

        public bool IsTestInstanceRunning
        {
            get
            {
                var runningProcesses = GetRunningAccessInstancesProcessId();
                return runningProcesses.Count(pid => pid == ProcessId) == 1;
            }
        }

        private static IList<int> GetDifference(ICollection<int> accessProcessesBefore, List<int> accessProcessesAfter)
        {
            return accessProcessesAfter.FindAll(pId => !accessProcessesBefore.Contains(pId));
        }

        private static List<int> GetRunningAccessInstancesProcessId()
        {
            var accessProcesses = Process.GetProcessesByName("MSACCESS");
            return GetProcessIds(accessProcesses);
        }

        private static List<int> GetProcessIds(ICollection<Process> accessProcesses)
        {
            var accessProcessIds = new List<int>(accessProcesses.Count);
            accessProcessIds.AddRange(accessProcesses.Select(accessProcess => accessProcess.Id));
            return accessProcessIds;
        }

        private string CreateTestDb()
        {
            var tempDbFileName = string.IsNullOrEmpty(_path) ? GetTempFileName() : _path;
            CreateTempDatabase(tempDbFileName);
            return tempDbFileName;
        }

        private void CreateTempDatabase(string tempDbFileName)
        {
            CreateTempDatabase(tempDbFileName, Locale);
        }

        private void CreateTempDatabase(string tempDbFileName, string locale)
        {
            var app = Application as Microsoft.Office.Interop.Access.Application;
            app.NewCurrentDatabase(tempDbFileName);
            app.CloseCurrentDatabase();
            /*
            using (var db = new ComWrapper<Database>(
                DBEngine.CreateDatabase(tempDbFileName, locale)))
            {
                db.ComReference.Close();
            }
            */
        }

        private static string GetTempFileName()
        {
            var tempFileName = Path.GetTempFileName();
            File.Delete(tempFileName);

            return tempFileName;
        }

        private void Cleanup()
        {
            if (!_isDisposed)
            {
                CleanupUnmanagedResources();
            }
            _isDisposed = true;
        }

        private void CleanupUnmanagedResources()
        {
            try
            {
                CloseCurrentDatabase();
            }
            catch (TimeoutException)
            {
                KillAccessInstance();
                Application = null;
            }
            finally
            {
                if (_dbEngine != null)
                    Marshal.ReleaseComObject(_dbEngine);
                ShutDownAccess();
                DeleteDatabaseFile();
            }
        }

        private void DeleteDatabaseFile()
        {
            if (!string.IsNullOrEmpty(TempDbFileName))
            {
                DeleteFileWithMultipleAttempts(TempDbFileName);
            }
        }

        private void ShutDownAccess()
        {
            if (Application is null)
                return;

            try
            {
                Application.Quit(AcQuitOption.acQuitSaveNone);
                Marshal.ReleaseComObject(Application);
            }
            catch (Exception xcp)
            {
                Trace.WriteLine(xcp);
                KillAccessInstance();
            }
            finally
            {
                Application = null;
            }
        }

        private void CloseCurrentDatabase()
        {
            Database currentDb = null;

            try
            {
                currentDb = (Database)Application.CurrentDb();
                if (currentDb != null)
                {
                    //CloseCurrentDatabaseWithTimeout();
                    Application.CloseCurrentDatabase();
                }
            }
            catch { }
            finally
            {
                if (currentDb != null)
                    Marshal.ReleaseComObject(currentDb);
            }
        }

        private void DeleteFileWithMultipleAttempts(string dbFileName)
        {
            var attempt = 0;
            bool couldDelete;
            do
            {
                attempt++;
                couldDelete = DeleteFile(dbFileName);
            } while (!couldDelete && attempt < MaxAttempt);
            if (couldDelete)
                Console.WriteLine("Deleted on {0}. attempt.", attempt);
            else
                Console.WriteLine("Could not delete in " + attempt + " attempts.");
        }

        private bool DeleteFile(string dbFileName)
        {
            try
            {
                File.Delete(dbFileName);
                return true;
            }
            catch (Exception xcp)
            {
                Trace.WriteLine(xcp);
                return false;
            }
        }

        private void KillAccessInstance()
        {
            Debug.Assert(ProcessId != 0, "********** Cannot kill the process because I don't know its ID.");

            var accessProcess = GetProcessByIdOrDefault(ProcessId);
            if (accessProcess is null)
            {
                Trace.WriteLine(
                    string.Format("********** This is thread \"{0}\". Cannot kill Access with process id #{1} because there is no process with this id.",
                        Thread.CurrentThread.Name, ProcessId));
                return;
            }
            if (accessProcess.HasExited)
            {
                Trace.WriteLine(
                    string.Format(
                        "********** This is thread \"{0}\". Cannot kill Access with process id #{1} because it has already exited.",
                        Thread.CurrentThread.Name, ProcessId));
                return;
            }

            try
            {
                Trace.WriteLine(
                    string.Format(
                        "********** This is thread \"{0}\". I am going to kill Access with process id #{1}",
                        Thread.CurrentThread.Name, ProcessId));
                accessProcess.Kill();
                accessProcess.WaitForExit();
                Trace.WriteLine(
                    string.Format(
                        "********** This is thread \"{0}\". Successfully killed Access with process id #{1}",
                        Thread.CurrentThread.Name, ProcessId));
            }
            catch (Exception xcp)
            {
                Trace.WriteLine(
                    string.Format(
                        "********** This is thread \"{0}\". There was an error while killing Access with process id #{1}:\r\n{2}\r\nStackTrace:\r\n{3}",
                        Thread.CurrentThread.Name, ProcessId, xcp.Message, xcp.StackTrace));
                throw;
            }
        }

        private Process GetProcessByIdOrDefault(int processId)
        {
            return Process.GetProcesses().FirstOrDefault(p => p.Id == processId);
        }

        /*
        private string GetLockFileName(string dbFileName)
        {
            /// @todo: Get lock-file (if existing)
            //throw new NotImplementedException();

            return null;
        }
        */

        private void CloseCurrentDatabaseWithTimeout()
        {
            var threadStart = new ParameterizedThreadStart(CloseCurrentDatabaseWorker);
            var closerThread = new Thread(threadStart) { Name = "Closer-Thread" };
            closerThread.Start(Application);

            var killerKilledWithinTime = closerThread.Join(DatabaseDeleteTimeout);

            Debug.WriteLine(string.Format("After Join({0})", killerKilledWithinTime ? "on time" : "after timeout"));

            if (!killerKilledWithinTime)
            {
                Trace.WriteLine("Before throwing TimeoutException");
                throw new TimeoutException(string.Format("Could not close the current database within {0}ms.", DatabaseDeleteTimeout));
            }
        }

        private static void CloseCurrentDatabaseWorker(object param)
        {
            var accessApplication = (dynamic)param;
            try
            {
                accessApplication.CloseCurrentDatabase();
            }
            catch (COMException)
            {
                // To avoid test run error with MSTest test runner
            }
        }

        public void AddAccessFileReference()
        {
            var libDbFileName = CreateTestDb();
            Application.References.AddFromFile(libDbFileName);
            var appdbFileName = Application.CurrentDb().Name;
            Application.CloseCurrentDatabase();
            Application.OpenCurrentDatabase(appdbFileName);
        }

        public void RemoveAllVbComponents()
        {
            var components = ActiveVBProject.VBComponents;
            foreach (VBComponent c in components)
            {
                components.Remove(c);
            }
        }


        public void ShowAccessWindow()
        {
            Application.Visible = true;
        }

        public bool Quit()
        {
            try
            {
                Application.Quit(AcQuitOption.acQuitSaveNone);
                Thread.Sleep(200);
                return !IsTestInstanceRunning;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
