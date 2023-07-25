using NUnit.Framework;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{
    internal class AssertTests
    {
        [Test]
        public void EqualToTest()
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
        [Ignore("only to check NUnit behaviour")]
        public void NunitEqualToTest()
        {
            Assert.That("1", Is.EqualTo(1));
        }
    }
}
