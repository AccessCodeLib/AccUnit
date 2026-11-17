using System.IO;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Keeps a copy of a source file before an export writes over it (the file may contain work not yet imported).
    /// </summary>
    public static class SourceFileBackup
    {
        private const string BackupExtension = ".bak";

        /// <returns>Path of the backup, or null if the file does not exist.</returns>
        public static string BackupIfExists(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            var backupPath = filePath + BackupExtension;
            File.Copy(filePath, backupPath, true);
            return backupPath;
        }
    }
}
