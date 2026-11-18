using System.ComponentModel;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Opens and closes the file with the VBA project under test. The tool names are host neutral,
    /// so further hosts (e.g. Excel) can be added later without changing the interface.
    /// </summary>
    [McpServerToolType]
    public class SessionTools
    {
        private readonly AccessSession _session;

        public SessionTools(AccessSession session)
        {
            _session = session;
        }

        [McpServerTool(Name = "accunit_open")]
        [Description("Opens the file with the VBA project under test in a visible Office instance owned by this server. " +
                     "Currently supported: Microsoft Access databases (.accdb, .accda, .mdb), opened exclusively. " +
                     "Only one file can be open at a time. The file should be in a trusted location, otherwise VBA code does not run. " +
                     "By default the startup code (AutoExec, startup form) is skipped like opening with the Shift key held down.")]
        public SessionStatus Open(
            [Description("Absolute path of the file (.accdb, .accda or .mdb).")] string filePath,
            [Description("Skip AutoExec and the startup form (Shift bypass). Set to false if the tests need the startup code. " +
                         "Has no effect if the database sets AllowBypassKey = False.")] bool bypassStartup = true)
        {
            try
            {
                return _session.Open(filePath, bypassStartup);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_close")]
        [Description("Closes the open file without saving and quits the Office instance of this server.")]
        public CloseResult Close()
        {
            return _session.Close();
        }

        [McpServerTool(Name = "accunit_status", ReadOnly = true)]
        [Description("Returns whether a file is open, its path, the application version, the process id and whether the file is trusted.")]
        public SessionStatus GetStatus()
        {
            return _session.GetStatus();
        }
    }
}
