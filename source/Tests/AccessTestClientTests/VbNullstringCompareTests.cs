using AccessCodeLib.Common.TestHelpers.AccessRelated;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
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
        [Ignore("Does not work, because vbNullstring becomes null")]
        public void CompareVbNullstringWithEmptyString()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test() as String
   Test = vbNullString      
End Function
");
            object ret = _accessTestHelper.Application.Run("Test");
            Type retType = ret.GetType();
            Assert.That(retType.FullName, Is.EqualTo("System.String"));
            Assert.That(ret, Is.EqualTo(""));
        }

        [Test]
        [Ignore("Does not work, because vbNullstring becomes null")]
        public void CompareVbNullstringWithEmptyString_WithRefParam()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test(ByRef ReturnValue as String)
   ReturnValue = vbNullString      
End Function
");
            string ret = "abc";
            _accessTestHelper.Application.Run("Test", ref ret);
            Assert.That(ret, Is.Not.EqualTo(null));

            Type retType = ret.GetType();
            Assert.That(retType.FullName, Is.EqualTo("System.String"));
            Assert.That(ret, Is.EqualTo(""));
        }


        [Test]
        public void AssertTest()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test(ByRef ReturnValue as String)
   ReturnValue = """"      
End Function
");
            string ret = "abc";
            _accessTestHelper.Application.Run("Test", ref ret);

            var assert = new Interop.Assert();
            assert.AreEqual(ret, "");
        }
    }
}
