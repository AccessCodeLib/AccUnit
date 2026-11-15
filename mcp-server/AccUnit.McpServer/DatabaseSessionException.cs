using System;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Error with a message that is safe to show to the MCP client.
    /// </summary>
    public class DatabaseSessionException : Exception
    {
        public DatabaseSessionException(string message) : base(message)
        {
        }

        public DatabaseSessionException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
