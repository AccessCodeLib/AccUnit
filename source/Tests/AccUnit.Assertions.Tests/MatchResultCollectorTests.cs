using NUnit.Framework;

namespace AccessCodeLib.AccUnit.Assertions.Tests
{

    public class MatchResultCollectorTests
    {
        [Test]
        public void ThatFillCollector()
        {
            var testCollector = new TestCollector();
            var assert = new Assertions
            {
                MatchResultCollector = testCollector
            };
            var Iz = new ConstraintBuilder();

            var actual = 1;
            var expected = 0;
            assert.That(actual, Iz.EqualTo(expected));

            Assert.That(testCollector.Result.Match, Is.EqualTo(false), testCollector.Result.Text);
        }
    }
}
