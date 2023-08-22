using AccessCodeLib.Common.TestHelpers.AccessRelated;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class TestSuiteTests
    {
        private AccessTestHelper _accessTestHelper;
        private Interop.ITestBuilder _testBuilder;

        [SetUp]
        public void TestBuilderTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper();
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
        public void CallByClassName_Select2MethodsAsEnumerable_CheckSummary()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

public Sub TestMethod1()
   m_Check = m_Check + 1
End Sub

public Sub TestMethod2()
   m_Check = m_Check + 2
End Sub

public Sub TestMethod3()
   m_Check = m_Check + 4
End Sub

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var testSuite = new AccessTestSuite
            {
                HostApplication = _accessTestHelper.Application
            };

            var methods = new string[] { "TestMethod1", "TestMethod3" };
            var summary = testSuite.AddByClassName("clsAccUnitTestClass").Select(methods).Run().Summary;

            Assert.That(summary.Passed, Is.EqualTo(2));
        }

        [Test]
        public void CallByClassNameInInteropAccessTestSuite_Select2MethodsAsString_CheckSummary()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

public Sub TestMethod1()
   m_Check = m_Check + 1
End Sub

public Sub TestMethod2()
   m_Check = m_Check + 2
End Sub

public Sub TestMethod3()
   m_Check = m_Check + 4
End Sub

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var testSuite = new Interop.AccessTestSuite
            {
                HostApplication = _accessTestHelper.Application
            };

            var testNameFilter = "TestMethod[13]";
            var summary = testSuite.AddByClassName("clsAccUnitTestClass").SelectTests(testNameFilter).Run().Summary;

            Assert.That(summary.Passed, Is.EqualTo(2));
        }

    }
}
