using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.TestHelpers.AccessRelated;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    // vbNullString is passed from VBA to C# as a BSTR null pointer, which becomes null (not "").
    // These tests document that behavior and show how Iz.StringCompare(vbNullStringEqualEmptyString) handles it.
    // The option matrix itself is tested in AccUnit.Assertions.Tests (InteropConstraintTests).
    [TestFixture]
    internal class VbNullstringCompareTests
    {
        private AccessTestHelper _accessTestHelper;

        [SetUp]
        public void AccessClientTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper();
        }

        [TearDown]
        public void AccessClientTestsCleanup()
        {
            _accessTestHelper?.Dispose();
            _accessTestHelper = null;
        }

        private static bool Matches(object actual, StringCompareMode compareMode, string expected, out string text)
        {
            var testCollector = new TestCollector();
            var assert = new Assertions.Assertions
            {
                MatchResultCollector = testCollector
            };
            var Iz = new Interop.ConstraintBuilder();
            assert.That(actual, Iz.StringCompare(compareMode).EqualTo(expected));
            text = testCollector.Result.Text;
            return testCollector.Result.Match;
        }

        private static void AssertVbNullStringEqualsEmptyStringOnlyWithOption(object actual)
        {
            Assert.That(actual, Is.Null);

            Assert.That(Matches(actual, StringCompareMode.BinaryCompare, "", out var text), Is.False, text);
            Assert.That(Matches(actual, StringCompareMode.BinaryCompare | StringCompareMode.vbNullStringEqualEmptyString, "", out text), Is.True, text);
        }

        [Test]
        public void CompareVbNullStringWithNull()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test() as String
   Test = vbNullString      
End Function
");
            object ret = _accessTestHelper.Application.Run("Test");
            Assert.That(ret, Is.EqualTo(null));
        }

        [Test]
        public void CompareVbNullstringWithEmptyString_RequiresStringCompareOption()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test() as String
   Test = vbNullString      
End Function
");
            object ret = _accessTestHelper.Application.Run("Test");
            AssertVbNullStringEqualsEmptyStringOnlyWithOption(ret);
        }

        [Test]
        public void CompareVbNullstringWithEmptyString_WithRefParam_RequiresStringCompareOption()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test(ByRef ReturnValue as String)
   ReturnValue = vbNullString      
End Function
");
            string ret = "abc";
            _accessTestHelper.Application.Run("Test", ref ret);
            AssertVbNullStringEqualsEmptyStringOnlyWithOption(ret);
        }

    }
}