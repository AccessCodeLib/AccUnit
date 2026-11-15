using System.ComponentModel;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AccessCodeLib.AccUnit.McpServer
{
    [McpServerToolType]
    public class SessionTools
    {
        private readonly AccessSession _session;

        public SessionTools(AccessSession session)
        {
            _session = session;
        }

        [McpServerTool(Name = "accunit_open_database")]
        [Description("Starts a hidden Access instance owned by this server and opens the database exclusively. " +
                     "Only one database can be open at a time. The database should be in a trusted location, otherwise VBA code does not run.")]
        public SessionStatus OpenDatabase(
            [Description("Absolute path of the .accdb, .accda or .mdb file.")] string databasePath)
        {
            try
            {
                return _session.Open(databasePath);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }

        [McpServerTool(Name = "accunit_close_database")]
        [Description("Closes the open database without saving and quits the Access instance of this server.")]
        public CloseResult CloseDatabase()
        {
            return _session.Close();
        }

        [McpServerTool(Name = "accunit_status", ReadOnly = true)]
        [Description("Returns whether a database is open, its path, the Access version, the process id and whether the database is trusted.")]
        public SessionStatus GetStatus()
        {
            return _session.GetStatus();
        }
    }
}
