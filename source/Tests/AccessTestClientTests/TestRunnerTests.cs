using AccessCodeLib.AccUnit.Integration;
using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System.Linq;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class TestRunnerTests
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
            if (_testBuilder != null)
            {
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
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var testRunner = new Interop.TestRunner();
            testRunner.Run(fixture, "TestMethod");

        }

        [Test]
        public void FindAndRunTestMethodesFromTestClass()
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
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner();
            testRunner.Run(fixture, "*", result);

            var actual = result.Results.Count();

            Assert.That(actual, Is.EqualTo(2));
            
        }

        [Test]
        public void FindAndRunTestMethodesFromTestClassWithSetupAndTeardown()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Value as Long

public sub Setup()
   m_Value = 123
End sub

public sub Teardown()
   m_Value = 1
End sub

public Function TestMethod1() as Long
   TestMethod1 = m_Value      
End Function
public Function TestMethod2() as Long
   Dim a as Long
   a = m_Value
End Function
private Function TestMethod3() as Long
   TestMethod3 = 999      
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner();
            testRunner.Run(fixture, "*", result);

            var resultCount = result.Results.Count();
            
            Assert.That(resultCount, Is.EqualTo(2));

            foreach (var testResult in result.Results)
            {
                var res = testResult as TestResult;
                Assert.That(res.IsSuccess, Is.EqualTo(true), res.Message);
            }
           

            var invocHelper = new InvocationHelper(fixture);
            var ValueAfterTeardowns = invocHelper.InvokeMethod("TestMethod1");
            Assert.That(ValueAfterTeardowns, Is.EqualTo(1));

        }
    }
}
