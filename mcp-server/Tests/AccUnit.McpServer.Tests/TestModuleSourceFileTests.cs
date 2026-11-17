using System;
using System.IO;
using System.Text;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class TestModuleSourceFileTests
    {
        private const string TestClassCode = "Option Compare Text\r\nOption Explicit\r\n\r\n'AccUnit:TestClass\r\n\r\n" +
                                             "Public Sub Test_Works()\r\n    Assert.That 1, Iz.EqualTo(1)\r\nEnd Sub\r\n";
        private const string ClassExportHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n" +
                                                 "Attribute VB_Name = \"modStringsTests\"\r\n";

        private static readonly Encoding Ansi = Encoding.GetEncoding(1252);

        private string _directory;

        [SetUp]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "AccUnitTestModuleSourceFileTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_directory, true);
        }

        private TestModuleSourceFile CreateFile(string fileName, string content, Encoding encoding)
        {
            var path = Path.Combine(_directory, fileName);
            File.WriteAllText(path, content, encoding);
            return TestModuleSourceFile.Read(path, Ansi);
        }

        [Test]
        public void Decode_Utf8WithBom_RemovesBomAndKeepsUmlauts()
        {
            var bytes = new UTF8Encoding(true).GetPreamble();
            bytes = Concat(bytes, Encoding.UTF8.GetBytes("' Gr\u00F6\u00DFe"));

            Assert.That(TestModuleSourceFile.Decode(bytes, Ansi), Is.EqualTo("' Gr\u00F6\u00DFe"));
        }

        [Test]
        public void Decode_Utf8WithoutBom_KeepsUmlauts()
        {
            Assert.That(TestModuleSourceFile.Decode(Encoding.UTF8.GetBytes("' Gr\u00F6\u00DFe"), Ansi), Is.EqualTo("' Gr\u00F6\u00DFe"));
        }

        [Test]
        public void Decode_Ansi_KeepsUmlauts()
        {
            Assert.That(TestModuleSourceFile.Decode(Ansi.GetBytes("' Gr\u00F6\u00DFe"), Ansi), Is.EqualTo("' Gr\u00F6\u00DFe"));
        }

        [TestCase("'AccUnit:TestClass")]
        [TestCase("' AccUnit:TestRelated")]
        [TestCase("'accunit:testclass  ")]
        public void HasTestModuleHeader_MarkerInHeader_IsTrue(string marker)
        {
            Assert.That(TestModuleSourceFile.HasTestModuleHeader("Option Explicit\r\n" + marker + "\r\n\r\nPublic Sub A()\r\nEnd Sub"), Is.True);
        }

        [Test]
        public void HasTestModuleHeader_MarkerOnlyInProcedure_IsFalse()
        {
            Assert.That(TestModuleSourceFile.HasTestModuleHeader("Option Explicit\r\n\r\nPublic Sub A()\r\n'AccUnit:TestClass\r\nEnd Sub"), Is.False);
        }

        [Test]
        public void HasTestModuleHeader_NoMarker_IsFalse()
        {
            Assert.That(TestModuleSourceFile.HasTestModuleHeader("Option Explicit\r\n\r\nPublic Function A()\r\nEnd Function"), Is.False);
        }

        [Test]
        public void GetRejectReason_TestClass_IsNull()
        {
            var file = CreateFile("modStringsTests.cls", TestClassCode, new UTF8Encoding(true));

            Assert.That(file.GetRejectReason(Ansi), Is.Null);
        }

        [Test]
        public void GetRejectReason_NoTestModule_IsReported()
        {
            var file = CreateFile("modStrings.bas", "Option Explicit\r\n\r\nPublic Function A()\r\nEnd Function\r\n", Ansi);

            Assert.That(file.GetRejectReason(Ansi), Does.Contain("Not a test module"));
        }

        [Test]
        public void GetRejectReason_VbNameDiffersFromFileName_IsReported()
        {
            var file = CreateFile("OtherName.cls", ClassExportHeader + TestClassCode, Ansi);

            Assert.That(file.GetRejectReason(Ansi), Does.Contain("file name must match the module name"));
        }

        [Test]
        public void GetRejectReason_CharacterNotInAnsiCodePage_IsReported()
        {
            var file = CreateFile("modStringsTests.cls", TestClassCode + "' \u4E2D\r\n", Encoding.UTF8);

            Assert.That(file.GetRejectReason(Ansi), Does.Contain("ANSI code page"));
        }

        [Test]
        public void GetImportSource_ClassWithoutExportHeader_AddsHeaderWithModuleName()
        {
            var file = CreateFile("modStringsTests.cls", TestClassCode, Encoding.UTF8);

            var source = file.GetImportSource();

            Assert.That(source, Does.StartWith("VERSION 1.0 CLASS\r\n"));
            Assert.That(source, Does.Contain("Attribute VB_Name = \"modStringsTests\"\r\n"));
            Assert.That(source, Does.EndWith(TestClassCode));
        }

        [Test]
        public void GetImportSource_ClassWithExportHeader_IsUnchanged()
        {
            var file = CreateFile("modStringsTests.cls", ClassExportHeader + TestClassCode, new UTF8Encoding(true));

            Assert.That(file.GetImportSource(), Is.EqualTo(ClassExportHeader + TestClassCode));
        }

        [Test]
        public void GetImportSource_StandardModuleWithoutVbName_AddsVbName()
        {
            var file = CreateFile("TestHelper.bas", "'AccUnit:TestRelated\r\nOption Explicit\r\n", Ansi);

            Assert.That(file.GetImportSource(), Does.StartWith("Attribute VB_Name = \"TestHelper\"\r\n'AccUnit:TestRelated"));
        }

        [Test]
        public void WriteAnsiCopy_Utf8WithLf_WritesAnsiWithCrLf()
        {
            var file = CreateFile("modStringsTests.cls", ClassExportHeader.Replace("\r\n", "\n") + "'AccUnit:TestClass\n' Gr\u00F6\u00DFe\n", new UTF8Encoding(true));
            var targetDirectory = Path.Combine(_directory, "ansi");
            Directory.CreateDirectory(targetDirectory);

            file.WriteAnsiCopy(targetDirectory, Ansi);

            var bytes = File.ReadAllBytes(Path.Combine(targetDirectory, "modStringsTests.cls"));
            Assert.That(bytes, Is.EqualTo(Ansi.GetBytes(ClassExportHeader + "'AccUnit:TestClass\r\n' Gr\u00F6\u00DFe\r\n")));
        }

        private static byte[] Concat(byte[] first, byte[] second)
        {
            var result = new byte[first.Length + second.Length];
            first.CopyTo(result, 0);
            second.CopyTo(result, first.Length);
            return result;
        }
    }
}
