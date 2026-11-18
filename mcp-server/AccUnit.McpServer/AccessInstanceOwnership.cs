using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Decides whether the Access instance returned by COM was started by this server.
    /// While another automation (e.g. another MCP server) uses Access, COM may hand out that running instance
    /// instead of starting a new one; the server must neither use nor quit it.
    /// </summary>
    public static class AccessInstanceOwnership
    {
        public const string ExistingInstanceMessage =
            "Access returned an instance that was already running (probably used by another automation), " +
            "so the database was not opened. Try again when the other Access automation has finished.";

        /// <param name="processId">Process of the instance returned by COM.</param>
        /// <param name="processIdsBeforeStart">Access processes that were running before the instance was requested.</param>
        public static bool IsStartedByServer(int processId, ICollection<int> processIdsBeforeStart)
        {
            return !processIdsBeforeStart.Contains(processId);
        }
    }
}
