using System;
using System.IO;
using System.Linq;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Checks a database path from the MCP client before Access is started.
    /// </summary>
    public static class DatabasePathValidator
    {
        // Only file types with readable VBA source; .accde/.mde contain compiled code only.
        private static readonly string[] AllowedExtensions = { ".accdb", ".accda", ".mdb" };

        private const string AccdbLockFileExtension = ".laccdb";
        private const string MdbLockFileExtension = ".ldb";
        private const string MdbExtension = ".mdb";

        /// <summary>
        /// Returns the normalized full path or throws a <see cref="DatabaseSessionException"/>.
        /// </summary>
        public static string Validate(string databasePath)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
                throw new DatabaseSessionException("The database path is empty.");

            var fullPath = GetAbsoluteFullPath(databasePath);
            if (fullPath == null)
                throw new DatabaseSessionException("The database path must be an absolute path.");

            var extension = Path.GetExtension(fullPath);
            if (!AllowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                throw new DatabaseSessionException(
                    "Unsupported file type '" + extension + "'. Allowed: " + string.Join(", ", AllowedExtensions) + ".");

            if (!File.Exists(fullPath))
                throw new DatabaseSessionException("The database file does not exist: " + fullPath);

            var lockFilePath = GetLockFilePath(fullPath);
            if (File.Exists(lockFilePath))
                throw new DatabaseSessionException(
                    "The database seems to be open in another Access instance (lock file found: " + lockFilePath + "). " +
                    "Close it there first. If no Access instance uses it, the lock file is stale and can be deleted.");

            return fullPath;
        }

        private static string GetAbsoluteFullPath(string path)
        {
            try
            {
                return Path.IsPathRooted(path) ? Path.GetFullPath(path) : null;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                throw new DatabaseSessionException("The database path is not a valid file path.", ex);
            }
        }

        public static string GetLockFilePath(string databasePath)
        {
            var lockExtension = string.Equals(Path.GetExtension(databasePath), MdbExtension, StringComparison.OrdinalIgnoreCase)
                ? MdbLockFileExtension
                : AccdbLockFileExtension;
            return Path.ChangeExtension(databasePath, lockExtension);
        }
    }
}
