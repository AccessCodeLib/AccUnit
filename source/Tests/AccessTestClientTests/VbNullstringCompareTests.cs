using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class VbNullstringCompareTests
    {
        private AccessTestHelper _accessTestHelper;
        private int i;

        [SetUp]
        public void AccessClientTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper(i++);
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
            var cm = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
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
            var cm = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test() as String
   Test = vbNullString      
End Function
");
            object ret = _accessTestHelper.Application.Run("Test");
            Type retType = ret.GetType();
            Assert.That(retType.FullName, Is.EqualTo("System.String"));
            Assert.That(ret, Is.EqualTo(""));
        }
    }
}
