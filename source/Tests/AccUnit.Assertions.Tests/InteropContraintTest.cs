using AccessCodeLib.AccUnit.Interfaces;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    public class InteropConstraintTests
    {
        [SetUp]
        public void Setup()
        {
            // Is.All
            // Is.EquivalentTo
            // Is.InRange

        }

        [Test]
        [TestCase(1, 1, true)]
        [TestCase(0, 0, true)]
        [TestCase(-1, -1, true)]
        [TestCase(1, 0, false)]
        [TestCase("abc", "abc", true)]
        [TestCase("abc", "xyz", false)]
        [TestCase("", "", true)]
        [TestCase("abc", "", false)]
        [TestCase(null, null, true)]
        [TestCase(1, null, false)]
        [TestCase(null, 1, false)]
        public void EqualTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new InteropTestCollector();
            var assert = new AccUnit.Interop.Assert
            {
                MatchResultCollector = testCollector
            };
            var Iz = new AccUnit.Interop.ConstraintBuilder();
            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        [TestCase("abc", "", StringCompareMode.TextCompare, false)]
        [TestCase("abc", "Abc", StringCompareMode.TextCompare, true)]
        [TestCase("abc", "Abc", StringCompareMode.BinaryCompare, false)]
        [TestCase("abc", "abc", StringCompareMode.TextCompare, true)]
        public void EqualTest_StringCompare(string actual, string expected, StringCompareMode ComparerMethod, bool expectedResult)
        {
            var testCollector = new InteropTestCollector();
            var assert = new AccUnit.Interop.Assert
            {
                MatchResultCollector = testCollector
            };
            var Iz = new AccUnit.Interop.ConstraintBuilder();
            assert.That(actual, Iz.StringCompare(ComparerMethod).EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }
    }
}
