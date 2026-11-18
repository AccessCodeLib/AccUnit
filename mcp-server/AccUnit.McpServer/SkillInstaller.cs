using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Installs the skill accunit-tests that is shipped with the server (folder skills next to the exe)
    /// into a project or the user profile. An installed skill is never overwritten without overwrite = true,
    /// because it may contain local improvements; overwritten files are backed up first.
    /// </summary>
    public sealed class SkillInstaller
    {
        public const string TargetClaude = "claude";
        public const string TargetAgents = "agents";
        public const string ScopeProject = "project";
        public const string ScopeUser = "user";

        private const string BackupExtension = ".bak";
        private static readonly Dictionary<string, string> TargetFolders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { TargetClaude, ".claude" },   // Claude Code (also read by Copilot)
            { TargetAgents, ".agents" }    // Codex, GitHub Copilot, Gemini CLI
        };

        private readonly string _bundledSkillDirectory;
        private readonly string _userProfileDirectory;

        public SkillInstaller()
            : this(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "skills", TestAuthoringGuide.SkillName),
                   Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
        {
        }

        public SkillInstaller(string bundledSkillDirectory, string userProfileDirectory)
        {
            _bundledSkillDirectory = bundledSkillDirectory;
            _userProfileDirectory = userProfileDirectory;
        }

        public string GetTargetDirectory(string target, string scope, string projectPath)
        {
            if (string.IsNullOrWhiteSpace(target) || !TargetFolders.TryGetValue(target, out var agentFolder))
                throw new DatabaseSessionException("Unknown target '" + target + "'. Allowed: " + TargetClaude + ", " + TargetAgents + ".");

            string baseDirectory;
            if (string.Equals(scope, ScopeUser, StringComparison.OrdinalIgnoreCase))
            {
                baseDirectory = _userProfileDirectory;
            }
            else if (string.IsNullOrWhiteSpace(scope) || string.Equals(scope, ScopeProject, StringComparison.OrdinalIgnoreCase))
            {
                baseDirectory = GetCheckedProjectDirectory(projectPath);
            }
            else
            {
                throw new DatabaseSessionException("Unknown scope '" + scope + "'. Allowed: " + ScopeProject + ", " + ScopeUser + ".");
            }

            return Path.Combine(baseDirectory, agentFolder, "skills", TestAuthoringGuide.SkillName);
        }

        public SkillInstallResult Install(string target, string scope, string projectPath, bool overwrite)
        {
            var sourceFiles = GetBundledSkillFiles();
            var targetDirectory = GetTargetDirectory(target, scope, projectPath);
            var result = new SkillInstallResult { TargetDirectory = targetDirectory };

            var existingFiles = sourceFiles.Where(file => File.Exists(GetTargetFile(file, targetDirectory))).ToList();
            var differingFiles = existingFiles.Where(file => !FilesAreEqual(file, GetTargetFile(file, targetDirectory))).ToList();

            if (existingFiles.Count == 0)
            {
                CopyFiles(sourceFiles, targetDirectory, result);
                result.Status = SkillInstallStatus.Installed;
            }
            else if (differingFiles.Count == 0 && existingFiles.Count == sourceFiles.Count)
            {
                result.Status = SkillInstallStatus.UpToDate;
            }
            else if (!overwrite)
            {
                result.Status = SkillInstallStatus.AlreadyInstalledDifferent;
                result.Message = "The installed skill differs from the version shipped with the server (it may contain local changes). " +
                                 "Nothing was changed; use overwrite = true to replace it (a backup is made).";
            }
            else
            {
                foreach (var file in differingFiles)
                {
                    var targetFile = GetTargetFile(file, targetDirectory);
                    var backupFile = targetFile + BackupExtension;
                    File.Copy(targetFile, backupFile, true);
                    result.BackupFiles.Add(backupFile);
                }
                CopyFiles(sourceFiles, targetDirectory, result);
                result.Status = SkillInstallStatus.Updated;
            }
            return result;
        }

        private List<string> GetBundledSkillFiles()
        {
            if (!Directory.Exists(_bundledSkillDirectory))
                throw new DatabaseSessionException("The skill shipped with the server was not found: " + _bundledSkillDirectory);
            return Directory.GetFiles(_bundledSkillDirectory, "*", SearchOption.AllDirectories).ToList();
        }

        private static string GetCheckedProjectDirectory(string projectPath)
        {
            if (string.IsNullOrWhiteSpace(projectPath))
                throw new DatabaseSessionException("projectPath is required for scope project (the root directory of the project).");
            if (!Path.IsPathRooted(projectPath))
                throw new DatabaseSessionException("projectPath must be an absolute path.");
            var fullPath = Path.GetFullPath(projectPath);
            if (!Directory.Exists(fullPath))
                throw new DatabaseSessionException("The project directory does not exist: " + fullPath);
            return fullPath;
        }

        private string GetTargetFile(string sourceFile, string targetDirectory)
        {
            var relativePath = sourceFile.Substring(_bundledSkillDirectory.TrimEnd(Path.DirectorySeparatorChar).Length + 1);
            return Path.Combine(targetDirectory, relativePath);
        }

        private void CopyFiles(IEnumerable<string> sourceFiles, string targetDirectory, SkillInstallResult result)
        {
            foreach (var file in sourceFiles)
            {
                var targetFile = GetTargetFile(file, targetDirectory);
                Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                File.Copy(file, targetFile, true);
                result.Files.Add(targetFile);
            }
        }

        private static bool FilesAreEqual(string file1, string file2)
        {
            return File.ReadAllBytes(file1).SequenceEqual(File.ReadAllBytes(file2));
        }
    }

    public enum SkillInstallStatus
    {
        Installed,
        UpToDate,
        AlreadyInstalledDifferent,
        Updated
    }

    public class SkillInstallResult
    {
        public SkillInstallStatus Status { get; set; }
        public string TargetDirectory { get; set; }
        public List<string> Files { get; set; } = new List<string>();
        public List<string> BackupFiles { get; set; } = new List<string>();
        public string Message { get; set; }
    }
}
