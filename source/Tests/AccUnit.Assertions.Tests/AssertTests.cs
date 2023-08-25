using NUnit.Framework;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    internal class AssertTests
    {
        [Test]
        public void EqualToTest_FailType()
        {
            var assert = new Assertions();
            var Iz = new ConstraintBuilder();

            var actual = "1";
            var expected = 1;

            Assert.Throws<Interfaces.AssertionException>(() =>
            {
                assert.That(actual, Iz.EqualTo(expected));
            });
        }

        [Test]
        public void EqualToTest_TestEmptyStrings()
        {
            var assert = new Assertions();
            var Iz = new ConstraintBuilder();

            object actual = "";
            object expected = "";

            assert.That(actual, Iz.EqualTo(expected));

        }

        [Test]
        public void EqualToTest_TestEmptyStrings_Interop()
        {
            var assert = new Interop.Assert();
            //var Iz = new ConstraintBuilder();

            object actual = "";
            object expected = "";

            assert.AreEqual(expected, actual);

        }

        [Test]
        public void EqualToTest_vbNullstringVsEmptyString_CompareAsEqual()
        {
            var assert = new Interop.Assert();
            var Iz = new StringConstraintBuilder(System.StringComparison.InvariantCulture, true);

            object actual = "";
            string expected = null;

            assert.That(actual, Iz.EqualTo(expected));  
        }
    }
}
