using System;
using System.IO;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class SkillInstallerTests
    {
        private string _root;
        private string _bundled;
        private string _userProfile;
        private string _project;
        private SkillInstaller _installer;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "AccUnitSkillInstallerTests", Guid.NewGuid().ToString("N"));
            _bundled = Path.Combine(_root, "server", "skills", "accunit-tests");
            _userProfile = Path.Combine(_root, "user");
            _project = Path.Combine(_root, "project");
            Directory.CreateDirectory(_bundled);
            Directory.CreateDirectory(_userProfile);
            Directory.CreateDirectory(_project);
            File.WriteAllText(Path.Combine(_bundled, "SKILL.md"), "shipped");
            _installer = new SkillInstaller(_bundled, _userProfile);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_root, true);
        }

        [TestCase("claude", ".claude")]
        [TestCase("agents", ".agents")]
        [TestCase("CLAUDE", ".claude")]
        public void GetTargetDirectory_ProjectScope_UsesAgentFolderInProject(string target, string agentFolder)
        {
            var directory = _installer.GetTargetDirectory(target, "project", _project);

            Assert.That(directory, Is.EqualTo(Path.Combine(_project, agentFolder, "skills", "accunit-tests")));
        }

        [Test]
        public void GetTargetDirectory_UserScope_UsesUserProfile()
        {
            var directory = _installer.GetTargetDirectory("agents", "user", null);

            Assert.That(directory, Is.EqualTo(Path.Combine(_userProfile, ".agents", "skills", "accunit-tests")));
        }

        [Test]
        public void GetTargetDirectory_DefaultScope_IsProject()
        {
            Assert.That(_installer.GetTargetDirectory("claude", null, _project), Does.StartWith(_project));
        }

        [TestCase("copilot", "project")]
        [TestCase("claude", "global")]
        public void GetTargetDirectory_UnknownTargetOrScope_Throws(string target, string scope)
        {
            Assert.Throws<DatabaseSessionException>(() => _installer.GetTargetDirectory(target, scope, _project));
        }

        [TestCase(null)]
        [TestCase("relative\\project")]
        public void GetTargetDirectory_ProjectScopeWithoutAbsolutePath_Throws(string projectPath)
        {
            Assert.Throws<DatabaseSessionException>(() => _installer.GetTargetDirectory("claude", "project", projectPath));
        }

        [Test]
        public void GetTargetDirectory_MissingProjectDirectory_Throws()
        {
            var missing = Path.Combine(_root, "missing");

            Assert.Throws<DatabaseSessionException>(() => _installer.GetTargetDirectory("claude", "project", missing));
        }

        [Test]
        public void Install_NotInstalled_CopiesSkill()
        {
            var result = _installer.Install("claude", "project", _project, false);

            Assert.That(result.Status, Is.EqualTo(SkillInstallStatus.Installed));
            Assert.That(File.ReadAllText(Path.Combine(result.TargetDirectory, "SKILL.md")), Is.EqualTo("shipped"));
        }

        [Test]
        public void Install_SameContent_IsUpToDate()
        {
            _installer.Install("claude", "project", _project, false);

            var result = _installer.Install("claude", "project", _project, false);

            Assert.That(result.Status, Is.EqualTo(SkillInstallStatus.UpToDate));
        }

        [Test]
        public void Install_LocallyChangedWithoutOverwrite_KeepsLocalChanges()
        {
            var first = _installer.Install("claude", "project", _project, false);
            var skillFile = Path.Combine(first.TargetDirectory, "SKILL.md");
            File.WriteAllText(skillFile, "local improvement");

            var result = _installer.Install("claude", "project", _project, false);

            Assert.That(result.Status, Is.EqualTo(SkillInstallStatus.AlreadyInstalledDifferent));
            Assert.That(File.ReadAllText(skillFile), Is.EqualTo("local improvement"));
        }

        [Test]
        public void Install_LocallyChangedWithOverwrite_ReplacesAndBacksUp()
        {
            var first = _installer.Install("claude", "project", _project, false);
            var skillFile = Path.Combine(first.TargetDirectory, "SKILL.md");
            File.WriteAllText(skillFile, "local improvement");

            var result = _installer.Install("claude", "project", _project, true);

            Assert.That(result.Status, Is.EqualTo(SkillInstallStatus.Updated));
            Assert.That(File.ReadAllText(skillFile), Is.EqualTo("shipped"));
            Assert.That(File.ReadAllText(skillFile + ".bak"), Is.EqualTo("local improvement"));
        }

        [Test]
        public void Install_BundledSkillMissing_Throws()
        {
            var installer = new SkillInstaller(Path.Combine(_root, "nothing"), _userProfile);

            Assert.Throws<DatabaseSessionException>(() => installer.Install("claude", "project", _project, false));
        }
    }
}
