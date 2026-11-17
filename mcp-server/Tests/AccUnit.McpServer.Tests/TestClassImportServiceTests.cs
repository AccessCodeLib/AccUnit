using System;
using System.IO;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class TestClassImportServiceTests
    {
        [TestCase("modStringsTests", "modStringsTests")]
        [TestCase(" modStringsTests ", "modStringsTests")]
        [TestCase("modStringsTests.cls", "modStringsTests")]
        [TestCase("TestHelper.BAS", "TestHelper")]
        [TestCase("*Tests", "*Tests")]
        public void NormalizeFileNameFilter_ValidFilter_ReturnsNameWithoutExtension(string filter, string expected)
        {
            Assert.That(TestClassImportService.NormalizeFileNameFilter(filter), Is.EqualTo(expected));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("  ")]
        [TestCase("..\\modStringsTests")]
        [TestCase("C:\\Test\\modStringsTests")]
        [TestCase("sub/modStringsTests")]
        public void NormalizeFileNameFilter_EmptyOrPath_Throws(string filter)
        {
            Assert.Throws<DatabaseSessionException>(() => TestClassImportService.NormalizeFileNameFilter(filter));
        }

        [Test]
        public void ValidateImportPath_ExistingAbsoluteFolder_ReturnsFullPath()
        {
            var folder = Path.GetTempPath();

            Assert.That(TestClassImportService.ValidateImportPath(folder), Is.EqualTo(Path.GetFullPath(folder)));
        }

        [Test]
        public void ValidateImportPath_RelativePath_Throws()
        {
            Assert.Throws<DatabaseSessionException>(() => TestClassImportService.ValidateImportPath("Tests\\NVdev"));
        }

        [Test]
        public void ValidateImportPath_MissingFolder_Throws()
        {
            var missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

            Assert.Throws<DatabaseSessionException>(() => TestClassImportService.ValidateImportPath(missing));
        }
    }
}
