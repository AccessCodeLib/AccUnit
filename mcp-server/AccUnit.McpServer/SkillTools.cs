using System.ComponentModel;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AccessCodeLib.AccUnit.McpServer
{
    [McpServerToolType]
    public class SkillTools
    {
        private readonly SkillInstaller _skillInstaller;

        public SkillTools(SkillInstaller skillInstaller)
        {
            _skillInstaller = skillInstaller;
        }

        [McpServerTool(Name = "accunit_install_skill")]
        [Description("Installs the skill accunit-tests (guide for writing and running AccUnit tests) shipped with this server. " +
                     "target \"claude\": .claude/skills (Claude Code, also read by GitHub Copilot); target \"agents\": .agents/skills " +
                     "(Codex, GitHub Copilot, Gemini CLI). scope \"project\" (default) installs into the project directory, so the skill is " +
                     "only active in projects that use AccUnit; scope \"user\" installs into the user profile. An installed skill is not " +
                     "overwritten unless overwrite = true (it may contain local improvements); changed files are backed up as .bak first.")]
        public SkillInstallResult InstallSkill(
            [Description("\"claude\" or \"agents\".")] string target = SkillInstaller.TargetClaude,
            [Description("\"project\" (default) or \"user\".")] string scope = SkillInstaller.ScopeProject,
            [Description("Absolute path of the project root directory (required for scope \"project\").")] string projectPath = null,
            [Description("Replace an installed skill that differs from the shipped version (backup as .bak).")] bool overwrite = false)
        {
            try
            {
                return _skillInstaller.Install(target, scope, projectPath, overwrite);
            }
            catch (DatabaseSessionException ex)
            {
                throw new McpException(ex.Message);
            }
        }
    }
}
