using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Access = Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Owns exactly one Access instance started by this server and the database opened in it.
    /// All COM calls run on the <see cref="StaComDispatcher"/>; the server never attaches to
    /// an Access instance it did not start.
    /// </summary>
    public sealed class AccessSession : IDisposable
    {
        private const int QuitTimeoutMilliseconds = 10000;
        private const bool OpenExclusive = true;
        private const string AllowBypassKeyPropertyName = "AllowBypassKey";
        private const int DaoPropertyNotFoundHResult = unchecked((int)0x800A0CC6); // DAO error 3270

        private readonly StaComDispatcher _dispatcher;
        private Access.Application _application;
        private Process _accessProcess;
        private string _databasePath;

        public AccessSession(StaComDispatcher dispatcher)
        {
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public SessionStatus Open(string databasePath, bool bypassStartup)
        {
            var fullPath = DatabasePathValidator.Validate(databasePath);
            return _dispatcher.Invoke(() => OpenOnStaThread(fullPath, bypassStartup));
        }

        public CloseResult Close()
        {
            return _dispatcher.Invoke(() => CloseOnStaThread());
        }

        public SessionStatus GetStatus()
        {
            return _dispatcher.Invoke(() => GetStatusOnStaThread());
        }

        private SessionStatus OpenOnStaThread(string fullPath, bool bypassStartup)
        {
            if (IsOpen())
            {
                if (string.Equals(_databasePath, fullPath, StringComparison.OrdinalIgnoreCase))
                    return GetStatusOnStaThread();

                throw new DatabaseSessionException(
                    "Another database is already open: " + _databasePath + ". Close it first.");
            }

            ReleaseAccess();
            StartAccess();

            StartupBypassDecision bypass;
            try
            {
                bypass = StartupBypassDecision.Create(bypassStartup, bypassStartup ? ReadAllowBypassKey(fullPath) : null);
                OpenCurrentDatabase(fullPath, bypass.PressShiftKey);
            }
            catch (COMException ex)
            {
                QuitAccess();
                throw new DatabaseSessionException("Access could not open the database: " + ex.Message, ex);
            }
            _databasePath = fullPath;

            var status = GetStatusOnStaThread();
            status.StartupBypassAttempted = bypass.PressShiftKey;
            status.Warning = bypass.Warning;
            return status;
        }

        private void OpenCurrentDatabase(string fullPath, bool pressShiftKey)
        {
            using (pressShiftKey ? new ShiftKeyScope() : null)
            {
                _application.OpenCurrentDatabase(fullPath, OpenExclusive);
            }
        }

        /// <returns>The value of AllowBypassKey (true if the property is not set), or null if it could not be read.</returns>
        private bool? ReadAllowBypassKey(string fullPath)
        {
            // Uses the DAO engine of the own Access instance, so the bitness always matches.
            Dao._DBEngine dbEngine = null;
            Dao.Database database = null;
            Dao.Properties properties = null;
            Dao.Property property = null;
            try
            {
                dbEngine = _application.DBEngine;
                database = dbEngine.OpenDatabase(fullPath, false, true);
                properties = database.Properties;
                property = properties[AllowBypassKeyPropertyName];
                return Convert.ToBoolean(property.Value);
            }
            catch (COMException ex) when (ex.ErrorCode == DaoPropertyNotFoundHResult)
            {
                return true;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                ReleaseComObject(property);
                ReleaseComObject(properties);
                if (database != null)
                {
                    database.Close();
                    ReleaseComObject(database);
                }
                ReleaseComObject(dbEngine);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null)
                Marshal.ReleaseComObject(comObject);
        }

        private void StartAccess()
        {
            _application = new Access.Application();
            // Visible, so the user can see what happens and dialogs do not block unnoticed.
            _application.Visible = true;
            _accessProcess = GetProcess(_application);
        }

        private CloseResult CloseOnStaThread()
        {
            if (_application == null)
                return new CloseResult { WasOpen = false, Message = "No database is open." };

            var closedPath = _databasePath;
            var killed = QuitAccess();
            return new CloseResult
            {
                WasOpen = true,
                DatabasePath = closedPath,
                Message = killed
                    ? "Access did not quit in time and was terminated (no changes were saved)."
                    : "Database closed and Access quit."
            };
        }

        /// <returns>true if the Access process had to be terminated.</returns>
        private bool QuitAccess()
        {
            try
            {
                if (IsProcessAlive())
                {
                    ReleaseUnreferencedComObjects();
                    // No CloseCurrentDatabase: with unsaved modules it shows a modal save prompt
                    // and blocks the automation. Quit with acQuitSaveNone does not ask.
                    _application.Quit(Access.AcQuitOption.acQuitSaveNone);
                }
            }
            catch (COMException)
            {
                // Access is not reachable any more; the process check below decides what to do.
            }
            finally
            {
                ReleaseApplication();
            }

            ReleaseUnreferencedComObjects();
            return EnsureProcessEnded();
        }

        /// <summary>
        /// The AccUnit core does not release every COM object it obtains (VBComponents, CodeModules, ...).
        /// Access only quits when all references are gone, so collect the unreferenced wrappers.
        /// </summary>
        private static void ReleaseUnreferencedComObjects()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Runs work with the Access application of the open database on the STA thread.
        /// </summary>
        public T Execute<T>(Func<Access.Application, T> work)
        {
            return _dispatcher.Invoke(() =>
            {
                if (!IsOpen())
                    throw new DatabaseSessionException("No database is open. Call accunit_open_database first.");
                return work(_application);
            });
        }

        private bool EnsureProcessEnded()
        {
            var process = _accessProcess;
            _accessProcess = null;
            _databasePath = null;

            if (process == null)
                return false;

            using (process)
            {
                if (process.WaitForExit(QuitTimeoutMilliseconds))
                    return false;

                // Only the instance started by this server is terminated.
                process.Kill();
                process.WaitForExit(QuitTimeoutMilliseconds);
                return true;
            }
        }

        private void ReleaseAccess()
        {
            ReleaseApplication();
            if (_accessProcess != null)
            {
                _accessProcess.Dispose();
                _accessProcess = null;
            }
            _databasePath = null;
        }

        private void ReleaseApplication()
        {
            if (_application == null)
                return;

            Marshal.FinalReleaseComObject(_application);
            _application = null;
        }

        private SessionStatus GetStatusOnStaThread()
        {
            if (!IsOpen())
            {
                var lostInstance = _application != null;
                ReleaseAccess();
                return new SessionStatus
                {
                    IsOpen = false,
                    Message = lostInstance ? "The Access instance was closed outside of this server." : "No database is open."
                };
            }

            return new SessionStatus
            {
                IsOpen = true,
                DatabasePath = _databasePath,
                AccessVersion = _application.Version,
                ProcessId = _accessProcess.Id,
                IsTrusted = IsCurrentProjectTrusted()
            };
        }

        private bool IsCurrentProjectTrusted()
        {
            // Every COM object obtained from Access must be released, otherwise Access does not quit.
            var currentProject = _application.CurrentProject;
            try
            {
                return currentProject.IsTrusted;
            }
            finally
            {
                Marshal.ReleaseComObject(currentProject);
            }
        }

        private bool IsOpen()
        {
            return _application != null && _databasePath != null && IsProcessAlive();
        }

        private bool IsProcessAlive()
        {
            if (_accessProcess == null)
                return false;
            _accessProcess.Refresh();
            return !_accessProcess.HasExited;
        }

        private static Process GetProcess(Access.Application application)
        {
            int processId;
            GetWindowThreadProcessId(new IntPtr(application.hWndAccessApp()), out processId);
            return Process.GetProcessById(processId);
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        public void Dispose()
        {
            _dispatcher.Invoke(() =>
            {
                if (_application != null)
                    QuitAccess();
            });
        }
    }

    public class SessionStatus
    {
        public bool IsOpen { get; set; }
        public string DatabasePath { get; set; }
        public string AccessVersion { get; set; }
        public int? ProcessId { get; set; }
        public bool? IsTrusted { get; set; }
        public bool? StartupBypassAttempted { get; set; }
        public string Warning { get; set; }
        public string Message { get; set; }
    }

    public class CloseResult
    {
        public bool WasOpen { get; set; }
        public string DatabasePath { get; set; }
        public string Message { get; set; }
    }
}
