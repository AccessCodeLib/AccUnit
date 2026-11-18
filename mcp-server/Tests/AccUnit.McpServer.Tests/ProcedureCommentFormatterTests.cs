using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class ProcedureCommentFormatterTests
    {
        [TestCase(null)]
        [TestCase("")]
        [TestCase("\r\n\r\n")]
        public void Format_NoHeader_ReturnsNull(string header)
        {
            Assert.That(ProcedureCommentFormatter.Format(header), Is.Null);
        }

        [Test]
        public void Format_NaturalDocBlock_RemovesMarkersAndSeparators()
        {
            const string header =
                "'---------------------------------------------------------------------------------------\r\n" +
                "' Function: CompareVersions\r\n" +
                "'---------------------------------------------------------------------------------------\r\n" +
                "'\r\n" +
                "' Compares two version strings.\r\n" +
                "'\r\n" +
                "' Parameters:\r\n" +
                "'    Version1 - first version\r\n" +
                "'    Version2 - second version\r\n" +
                "'\r\n" +
                "' Returns:\r\n" +
                "'    -1, 0 or 1\r\n" +
                "'\r\n" +
                "'---------------------------------------------------------------------------------------\r\n";

            var comment = ProcedureCommentFormatter.Format(header);

            Assert.That(comment, Is.EqualTo(
                "Function: CompareVersions\r\n" +
                "\r\n" +
                "Compares two version strings.\r\n" +
                "\r\n" +
                "Parameters:\r\n" +
                "   Version1 - first version\r\n" +
                "   Version2 - second version\r\n" +
                "\r\n" +
                "Returns:\r\n" +
                "   -1, 0 or 1"));
        }

        [Test]
        public void Format_SimpleCommentWithBlankLines_KeepsTextOnly()
        {
            const string header = "\r\n\r\n' Returns the net price.\r\n\r\n";

            Assert.That(ProcedureCommentFormatter.Format(header), Is.EqualTo("Returns the net price."));
        }

        [Test]
        public void Format_RemComment_IsRecognized()
        {
            Assert.That(ProcedureCommentFormatter.Format("Rem Old style comment\r\n"), Is.EqualTo("Old style comment"));
        }

        [Test]
        public void Format_OnlySeparators_ReturnsNull()
        {
            Assert.That(ProcedureCommentFormatter.Format("'----------\r\n'==========\r\n"), Is.Null);
        }
    }
}
