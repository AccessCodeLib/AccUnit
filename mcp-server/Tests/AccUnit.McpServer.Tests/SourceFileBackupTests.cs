using System;
using System.IO;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class SourceFileBackupTests
    {
        private string _directory;

        [SetUp]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "AccUnitSourceFileBackupTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_directory, true);
        }

        [Test]
        public void BackupIfExists_MissingFile_ReturnsNull()
        {
            Assert.That(SourceFileBackup.BackupIfExists(Path.Combine(_directory, "modCalcTests.cls")), Is.Null);
        }

        [Test]
        public void BackupIfExists_ExistingFile_CopiesToBak()
        {
            var filePath = Path.Combine(_directory, "modCalcTests.cls");
            File.WriteAllText(filePath, "work in progress");

            var backupPath = SourceFileBackup.BackupIfExists(filePath);

            Assert.That(backupPath, Is.EqualTo(filePath + ".bak"));
            Assert.That(File.ReadAllText(backupPath), Is.EqualTo("work in progress"));
            Assert.That(File.Exists(filePath), Is.True);
        }

        [Test]
        public void BackupIfExists_BakExists_IsReplaced()
        {
            var filePath = Path.Combine(_directory, "modCalcTests.cls");
            File.WriteAllText(filePath, "new");
            File.WriteAllText(filePath + ".bak", "old");

            SourceFileBackup.BackupIfExists(filePath);

            Assert.That(File.ReadAllText(filePath + ".bak"), Is.EqualTo("new"));
        }
    }
}
