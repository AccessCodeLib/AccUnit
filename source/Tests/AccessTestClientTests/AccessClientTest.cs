using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System.Linq;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    public class AccessClientTests
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
            if (_accessTestHelper != null)
                _accessTestHelper.Dispose();
            _accessTestHelper = null;
        }

        [Test]
        public void AddAndRunFunction()
        {
            var cm = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function Test() as Long
   Test = 123      
End Function
");
            var ret = _accessTestHelper.Application.Run("Test");

            Assert.That(ret, Is.EqualTo(123));
        }

        [Test]
        public void AddAndRunClassMethod()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"  
public Function TestMethod() as Long
   TestMethod = 123      
End Function
");
            
            var cm = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitFactory", vbext_ComponentType.vbext_ct_StdModule, @"
public Function GetTestClassTestValue() as Long
   Dim testClass as clsAccUnitTestClass
   Set testClass = New clsAccUnitTestClass  
   GetTestClassTestValue = testClass.TestMethod
End Function
");
            var result = _accessTestHelper.Application.Run("GetTestClassTestValue");
            Assert.That(result, Is.EqualTo(123));
        }

        [Test]
        public void AddAndRunTestClass()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Long
   TestMethod = 123      
End Function
");
            var cm = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "AccUnit_TestClassFactory", vbext_ComponentType.vbext_ct_StdModule, @"
Public Function AccUnitTestClassFactory_clsAccUnitTestClass() As Object
   Set AccUnitTestClassFactory_clsAccUnitTestClass = New clsAccUnitTestClass
End Function
");
            var testBuilder = new Interop.TestBuilder();
            testBuilder.HostApplication = _accessTestHelper.Application;
            
            var fixture = testBuilder.CreateTest("clsAccUnitTestClass");

            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("TestMethod");
            
            testBuilder.Dispose();

            Assert.That(actual, Is.EqualTo(123));
        }
    }
}
