using AccessCodeLib.Common.VBIDETools;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace TestGeneratorTests
{
    public class VbaSourceFileTests
    {
        private static readonly Encoding Ansi1252 = Encoding.GetEncoding(1252);
        private const string GermanText = "' Größe ändern: äöüß €";
        private const string NonAnsiText = "' Pfeil → und 中";

        private string _directory;

        [SetUp]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "AccUnitVbaSourceFileTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_directory, true);
        }

        [Test]
        public void Decode_Utf8WithBom_ReturnsTextWithoutBom()
        {
            var bytes = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(GermanText)).ToArray();

            Assert.That(VbaSourceFile.Decode(bytes, Ansi1252), Is.EqualTo(GermanText));
        }

        [Test]
        public void Decode_Utf8WithoutBom_ReturnsText()
        {
            var bytes = Encoding.UTF8.GetBytes(GermanText);

            Assert.That(VbaSourceFile.Decode(bytes, Ansi1252), Is.EqualTo(GermanText));
        }

        [Test]
        public void Decode_Ansi_ReturnsText()
        {
            var bytes = Ansi1252.GetBytes(GermanText);

            Assert.That(VbaSourceFile.Decode(bytes, Ansi1252), Is.EqualTo(GermanText));
        }

        [TestCase(GermanText, true)]
        [TestCase(NonAnsiText, false)]
        public void CanBeEncoded_ReturnsExpected(string text, bool expected)
        {
            Assert.That(VbaSourceFile.CanBeEncoded(text, Ansi1252), Is.EqualTo(expected));
        }

        [Test]
        public void ToCrLf_MixedLineEndings_ReturnsCrLf()
        {
            Assert.That(VbaSourceFile.ToCrLf("a\nb\r\nc\n"), Is.EqualTo("a\r\nb\r\nc\r\n"));
        }

        [Test]
        public void WriteAnsiCopy_Utf8FileWithLf_WritesAnsiWithCrLf()
        {
            var sourceFile = Path.Combine(_directory, "Test1.cls");
            File.WriteAllText(sourceFile, "Option Explicit\n" + GermanText + "\n", new UTF8Encoding(true));
            var targetDirectory = Directory.CreateDirectory(Path.Combine(_directory, "ansi")).FullName;

            var ansiFile = VbaSourceFile.WriteAnsiCopy(sourceFile, targetDirectory, Ansi1252);

            Assert.That(ansiFile, Is.EqualTo(Path.Combine(targetDirectory, "Test1.cls")));
            Assert.That(File.ReadAllBytes(ansiFile), Is.EqualTo(Ansi1252.GetBytes("Option Explicit\r\n" + GermanText + "\r\n")));
        }

        [Test]
        public void WriteAnsiCopy_AnsiFile_WritesSameText()
        {
            var sourceFile = Path.Combine(_directory, "Test2.bas");
            var ansiBytes = Ansi1252.GetBytes("Option Explicit\r\n" + GermanText + "\r\n");
            File.WriteAllBytes(sourceFile, ansiBytes);
            var targetDirectory = Directory.CreateDirectory(Path.Combine(_directory, "ansi")).FullName;

            var ansiFile = VbaSourceFile.WriteAnsiCopy(sourceFile, targetDirectory, Ansi1252);

            Assert.That(File.ReadAllBytes(ansiFile), Is.EqualTo(ansiBytes));
        }

        [Test]
        public void WriteAnsiCopy_NonAnsiCharacters_ThrowsAndWritesNothing()
        {
            var sourceFile = Path.Combine(_directory, "Test3.cls");
            File.WriteAllText(sourceFile, NonAnsiText, new UTF8Encoding(true));
            var targetDirectory = Directory.CreateDirectory(Path.Combine(_directory, "ansi")).FullName;

            var ex = Assert.Throws<UnsupportedSourceCharactersException>(
                () => VbaSourceFile.WriteAnsiCopy(sourceFile, targetDirectory, Ansi1252));

            Assert.That(ex.FilePaths, Is.EqualTo(new[] { sourceFile }));
            Assert.That(Directory.GetFiles(targetDirectory), Is.Empty);
        }

        [Test]
        public void GetFilesWithUnsupportedCharacters_ReturnsOnlyRejectedFiles()
        {
            var ansiFile = Path.Combine(_directory, "Ansi.cls");
            File.WriteAllBytes(ansiFile, Ansi1252.GetBytes(GermanText));
            var utf8File = Path.Combine(_directory, "Utf8.cls");
            File.WriteAllText(utf8File, GermanText, new UTF8Encoding(false));
            var rejectedFile = Path.Combine(_directory, "Rejected.cls");
            File.WriteAllText(rejectedFile, NonAnsiText, new UTF8Encoding(true));

            var result = VbaSourceFile.GetFilesWithUnsupportedCharacters(new[] { ansiFile, utf8File, rejectedFile }, Ansi1252);

            Assert.That(result, Is.EqualTo(new[] { rejectedFile }));
        }

        [Test]
        public void ConvertExportedFile_AnsiFile_WritesUtf8WithBomAndCrLf()
        {
            var file = Path.Combine(_directory, "Exported.cls");
            File.WriteAllBytes(file, Ansi1252.GetBytes("Option Explicit\r\n" + GermanText + "\n"));

            VbaSourceFile.ConvertExportedFile(file, VbaSourceFile.Utf8BomEncoding, Ansi1252);

            var expected = new UTF8Encoding(true).GetPreamble()
                .Concat(Encoding.UTF8.GetBytes("Option Explicit\r\n" + GermanText + "\r\n")).ToArray();
            Assert.That(File.ReadAllBytes(file), Is.EqualTo(expected));
        }
    }
}
