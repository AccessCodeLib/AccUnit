using AccessCodeLib.AccUnit.TestRunner;
using NUnit.Framework;

namespace TestGeneratorTests
{
    public class TestNameFilterMatcherTests
    {
        [TestCase("TestMethod1", "TestMethod1", true)]
        [TestCase("TestMethod1", "TestMethod2", false)]
        [TestCase("Test*", "TestMethod1", true)]
        [TestCase("*Method1", "TestMethod1", true)]
        [TestCase("Te?t*", "TestMethod1", true)]
        [TestCase("Te?t*[13]", "TestMethod3", true)]
        [TestCase("Te?t*[13]", "TestMethod2", false)]
        [TestCase("TestMethod[1-3]", "TestMethod2", true)]
        [TestCase("TestMethod[!1]", "TestMethod1", false)]
        [TestCase("TestMethod[!1]", "TestMethod2", true)]
        [TestCase("TestMethod[^1]", "TestMethod2", true)]
        [TestCase("TestMethod[!1]", "TestMethod!", true)]
        public void IsMatch_WithSingleFilter_ReturnsExpected(string filter, string testName, bool expected)
        {
            var matcher = new TestNameFilterMatcher(new[] { filter });

            Assert.That(matcher.IsMatch(testName), Is.EqualTo(expected));
        }

        [TestCase("Test.Method", "Test.Method", true)]
        [TestCase("Test.Method", "TestXMethod", false)]
        [TestCase("Test+", "Test+", true)]
        [TestCase("Test+", "Testt", false)]
        [TestCase("Test(1)", "Test(1)", true)]
        [TestCase("Test(*", "Test(1", true)]
        [TestCase("Test[", "Test[", true)]
        [TestCase("Test[]", "Test[]", true)]
        [TestCase("Test[!]", "Test[!]", true)]
        [TestCase("A|B", "A", false)]
        [TestCase("A|B", "A|B", true)]
        [TestCase(@"A\d", @"A\d", true)]
        [TestCase(@"A\d", "A1", false)]
        [TestCase("^A$", "A", false)]
        public void IsMatch_WithSpecialCharacters_TreatsThemAsLiterals(string filter, string testName, bool expected)
        {
            var matcher = new TestNameFilterMatcher(new[] { filter });

            Assert.That(matcher.IsMatch(testName), Is.EqualTo(expected));
        }

        [TestCase("testmethod1", "TestMethod1")]
        [TestCase("TESTMETHOD*", "TestMethod1")]
        [TestCase("Test[a-z]ethod1", "TestMethod1")]
        public void IsMatch_WithDifferentCase_ReturnsTrue(string filter, string testName)
        {
            var matcher = new TestNameFilterMatcher(new[] { filter });

            Assert.That(matcher.IsMatch(testName), Is.True);
        }

        [TestCase("Foo1", true)]
        [TestCase("FooXYZ", true)]
        [TestCase("Bar", true)]
        [TestCase("XBar", false)]
        [TestCase("BarX", false)]
        [TestCase("Baz", false)]
        public void IsMatch_WithMultipleFilters_AnchorsApplyToEachFilter(string testName, bool expected)
        {
            var matcher = new TestNameFilterMatcher(new[] { "Foo*", "Bar" });

            Assert.That(matcher.IsMatch(testName), Is.EqualTo(expected));
        }

        [TestCase("Foo", false)]
        [TestCase("Foo*", true)]
        [TestCase("Fo?", true)]
        [TestCase("Foo[1]", true)]
        [TestCase("Foo.Bar", false)]
        [TestCase("", false)]
        [TestCase(null, false)]
        public void ContainsWildcard_ReturnsExpected(string filter, bool expected)
        {
            Assert.That(TestNameFilterMatcher.ContainsWildcard(filter), Is.EqualTo(expected));
        }
    }
}
