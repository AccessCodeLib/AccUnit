using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Office.Interop.Access;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    public class TestBuilderTests
    {
        private AccessTestHelper _accessTestHelper;
        private Interop.ITestBuilder _testBuilder;
        private int i;

        [SetUp]
        public void TestBuilderTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper(i++);
            _testBuilder = new Interop.TestBuilder
            {
                HostApplication = _accessTestHelper.Application
            };
        }

        [TearDown]
        public void TestBuilderTestsCleanup()
        {
            if (_testBuilder != null) {
                _testBuilder.Dispose();
                _testBuilder = null;
            }
            
            if (_accessTestHelper != null)
            {
                _accessTestHelper.Dispose();
                _accessTestHelper = null;
            }
        }
        
        [Test]
        public void CreateTestFromExistingFactoryMethode()
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
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);
        }

        [Test]
        public void CreateTestWithoutExistingFactoryMethode()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Long
   TestMethod = 123      
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod("TestMethod");
            
            Assert.That(actual, Is.EqualTo(123));
        }
    }
}
