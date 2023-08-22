using NUnit.Framework;
using System;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    public class ConstraintTests
    {
        [SetUp]
        public void Setup()
        {
            // Is.All
            // Is.EquivalentTo
            // Is.InRange

        }

        private static Assertions NewTestAssert(TestCollector testCollector, bool strict = false)
        {
            return new Assertions(strict)
            {
                MatchResultCollector = testCollector
            };
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
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
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
        [TestCase("1", 1, false)] // string ist not numeric
        [TestCase(1, "1", false)] // string ist not numeric
        public void AreEqualTest_InteropAssert(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);

            assert.AreEqual(expected, actual);
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        [TestCase(1, 1, false)]
        [TestCase(0, 0, false)]
        [TestCase(-1, -1, false)]
        [TestCase(1, 0, true)]
        [TestCase("abc", "abc", false)]
        [TestCase("abc", "xyz", true)]
        [TestCase("", "", false)]
        [TestCase("abc", "", true)]
        [TestCase("", "xyz", true)]
        [TestCase(1, null, true)]
        [TestCase(null, 1, true)]
        public void NotEqualTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.Not.Not.Not.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        [TestCase(1, 1, false)]
        [TestCase(0, 0, false)]
        [TestCase(-1, 0, true)]
        [TestCase(1, 0, false)]
        [TestCase(1, 2, true)]
        [TestCase("abc", "abc", false)]
        [TestCase("abc", "xyz", true)]
        [TestCase("", "", false)]
        [TestCase("abc", "", false)]
        [TestCase("", "xyz", true)]
        public void LessThanTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.LessThan(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult));
        }

        [Test]
        [TestCase(1, 1, true)]
        [TestCase(0, 0, true)]
        [TestCase(-1, -1, true)]
        [TestCase(1, 0, false)]
        [TestCase(1, 2, true)]
        [TestCase("abc", "abc", true)]
        [TestCase("abc", "xyz", true)]
        [TestCase("", "", true)]
        [TestCase("abc", "", false)]
        [TestCase("", "xyz", true)]
        public void LessThanOrEqualTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.LessThanOrEqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult));
        }

        [Test]
        [TestCase(1, 1, false)]
        [TestCase(0, 0, false)]
        [TestCase(-1, -1, false)]
        [TestCase(1, 0, true)]
        [TestCase(1, -1, true)]
        [TestCase("abc", "abc", false)]
        [TestCase("abc", "xyz", false)]
        [TestCase("", "", false)]
        [TestCase("abc", "", true)]
        [TestCase("", "xyz", false)]
        public void GreaterThanTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.GreaterThan(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        [TestCase(1, 1, true)]
        [TestCase(0, 0, true)]
        [TestCase(-1, -1, true)]
        [TestCase(1, 0, true)]
        [TestCase(1, -1, true)]
        [TestCase(0, -1, true)]
        [TestCase(-5, -1, false)]
        [TestCase(1, 2, false)]
        [TestCase("abc", "abc", true)]
        [TestCase("abc", "xyz", false)]
        [TestCase("", "", true)]
        [TestCase("abc", "", true)]
        [TestCase("", "xyz", false)]
        public void GreaterThanOrEqualTest(object actual, object expected, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.GreaterThanOrEqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        [TestCase(1, false)]
        [TestCase(null, true)]
        public void IsNullTest(object actual, bool expectedResult)
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(actual, Iz.Null);
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(expectedResult), result.Text);
        }

        [Test]
        public void DbNullIsNullTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(DBNull.Value, Iz.Null);
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(false), result.Text);
        }

        [Test]
        public void DbNullIsDBNullTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(DBNull.Value, Iz.DBNull);
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void DbNullEqualToDBNullTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(DBNull.Value, Iz.EqualTo(DBNull.Value));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void DbNullEqualToNumericTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            assert.That(DBNull.Value, Iz.EqualTo(1));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(false), result.Text);
        }

        [Test]
        public void IntArrayIsEqualTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            int[] expected = new int[] { 1, 2, 3 };
            int[] actual = new int[] { 1, 2, 3 };

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void StringArrayIsEqualTest()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            string[] expected = new string[] { "a", "b", "c" };
            string[] actual = new string[] { "a", "b", "c" };

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }

        [Test]
        public void IntArrayIsNotEqual_DifferentLength()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder();

            int[] expected = new int[] { 1, 2, 3 };
            int[] actual = new int[] { 1, 2 };

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(false), result.Text);
        }

        [Test]
        public void StrictConstraintBuilder_Int32IsNotDouble()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new ConstraintBuilder(true);

            int expected = 1;
            double actual = 1;

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(false), result.Text);
        }

        [Test]
        public void StrictAssertEqual_Int32IsNotDouble()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector, true);

            int expected = 1;
            double actual = 1;

            assert.AreEqual(expected, actual);
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(false), result.Text);
        }

        [Test]
        public void StringConstraintBuilder_StringCompareAisa()
        {
            var testCollector = new TestCollector();
            var assert = NewTestAssert(testCollector);
            var Iz = new StringConstraintBuilder(StringComparison.InvariantCultureIgnoreCase);

            string expected = "A";
            string actual = "a";

            assert.That(actual, Iz.EqualTo(expected));
            var result = testCollector.Result;

            Assert.That(result.Match, Is.EqualTo(true), result.Text);
        }



    }
}