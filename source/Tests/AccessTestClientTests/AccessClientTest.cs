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


        [Test]
        [Ignore("don't use TLBINF32.dll")]
        public void FindMethodeNameWithTLI()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Long
   TestMethod = 123      
End Function
");
            var testBuilder = new Interop.TestBuilder();
            testBuilder.HostApplication = _accessTestHelper.Application;

            var fixture = testBuilder.CreateTest("clsAccUnitTestClass");

            Assert.That(fixture, Is.Not.Null);

            var members = AccessCodeLib.Common.VBIDETools.TypeLib.TypeLibTools.GetTLIInterfaceMemberNames(fixture);
            Assert.That(members, Is.Not.Null);
            Assert.That(members.Count, Is.EqualTo(1));

            var member = members.ElementAt(0);
            Assert.That(member, Is.EqualTo("TestMethod"));

            var invocHelper = new InvocationHelper(fixture);
            var actual = invocHelper.InvokeMethod(member);

            testBuilder.Dispose();

            Assert.That(actual, Is.EqualTo(123));
        }

        [Test]
        [Ignore("don't use TLBINF32.dll")]
        public void FindPublicMethodesWithTLI()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod1() as Long
   TestMethod1 = 123      
End Function

public Function TestMethod2() as Long
   TestMethod2 = 123      
End Function

private Function TestMethod3() as Long
   TestMethod3 = 999      
End Function
");
            object fixture;
            using (var testBuilder = new Interop.TestBuilder())
            {
                testBuilder.HostApplication = _accessTestHelper.Application;
                fixture = testBuilder.CreateTest("clsAccUnitTestClass");
            }
            
            Assert.That(fixture, Is.Not.Null);

            var members = AccessCodeLib.Common.VBIDETools.TypeLib.TypeLibTools.GetTLIInterfaceMemberNames(fixture);
            Assert.That(members, Is.Not.Null);
            Assert.That(members.Count, Is.EqualTo(2));

            var invocHelper = new InvocationHelper(fixture);

            foreach (var member in members)
            {
                var actual = invocHelper.InvokeMethod(member);
                Assert.That(actual, Is.EqualTo(123));
            }
        }
    }
}
