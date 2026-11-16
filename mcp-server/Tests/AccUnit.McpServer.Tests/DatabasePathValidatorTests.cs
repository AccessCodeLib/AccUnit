using System;
using System.IO;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class DatabasePathValidatorTests
    {
        private string _directory;

        [SetUp]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "AccUnitMcpServerTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_directory, true);
        }

        [TestCase("test.accdb")]
        [TestCase("test.ACCDB")]
        [TestCase("test.accda")]
        [TestCase("test.mdb")]
        public void Validate_ExistingSupportedFile_ReturnsFullPath(string fileName)
        {
            var path = CreateFile(fileName);

            Assert.That(DatabasePathValidator.Validate(path), Is.EqualTo(path));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void Validate_EmptyPath_Throws(string path)
        {
            Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate(path));
        }

        [Test]
        public void Validate_RelativePath_Throws()
        {
            var ex = Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate(@"folder\test.accdb"));
            Assert.That(ex.Message, Does.Contain("absolute"));
        }

        [TestCase("test.accde")]
        [TestCase("test.mde")]
        [TestCase("test.txt")]
        public void Validate_UnsupportedExtension_Throws(string fileName)
        {
            var path = CreateFile(fileName);

            var ex = Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate(path));
            Assert.That(ex.Message, Does.Contain("Unsupported file type"));
        }

        [Test]
        public void Validate_MissingFile_Throws()
        {
            var path = Path.Combine(_directory, "missing.accdb");

            var ex = Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate(path));
            Assert.That(ex.Message, Does.Contain("does not exist"));
        }

        [TestCase("test.accdb", "test.laccdb")]
        [TestCase("test.mdb", "test.ldb")]
        public void Validate_LockFileExists_Throws(string fileName, string lockFileName)
        {
            var path = CreateFile(fileName);
            CreateFile(lockFileName);

            var ex = Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate(path));
            Assert.That(ex.Message, Does.Contain("lock file"));
        }

        [Test]
        public void Validate_InvalidCharacters_Throws()
        {
            Assert.Throws<DatabaseSessionException>(() => DatabasePathValidator.Validate("C:\\in<valid>.accdb"));
        }

        private string CreateFile(string fileName)
        {
            var path = Path.Combine(_directory, fileName);
            File.WriteAllText(path, string.Empty);
            return path;
        }
    }
}
