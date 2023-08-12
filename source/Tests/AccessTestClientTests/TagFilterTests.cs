using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using System.Linq;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class TagFilterTests
    {
        private AccessTestHelper _accessTestHelper;
        private Interop.ITestBuilder _testBuilder;

        [SetUp]
        public void AccessClientTestsSetup()
        {
            _accessTestHelper = AccessClientTestHelper.NewAccessTestHelper();
            _testBuilder = new Interop.TestBuilder
            {
                HostApplication = _accessTestHelper.Application
            };
        }

        [TearDown]
        public void AccessClientTestsCleanup()
        {
            _testBuilder?.Dispose();
            _testBuilder = null;

            _accessTestHelper?.Dispose();
            _accessTestHelper = null;
        }

        [Test]
        public void RunRowTest_runOnlyWithTagABC_CheckTagAndCheckValue2()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(1)
'AccUnit:Row(2).Tags(""ABC"")
'AccUnit:Row(3)
public Function TestMethod1(ByVal x as Long) as Long
   m_Check = x
   TestMethod1 = x
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var memberName = "TestMethod1";
            
            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[1].Tags.First().Name, Is.EqualTo("ABC"));

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result, "ABC");

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }

        [Test]
        public void RunRowTest_runOnlyWithTagABC_CheckTagAndCheckValueSum6()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Tags(""ABC"")
'AccUnit:Row(1)
'AccUnit:Row(2).Tags(""XYZ"")
'AccUnit:Row(3)
public Function TestMethod1(ByVal x as Long) as Long
   m_Check = m_Check + x
   TestMethod1 = x
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var memberName = "TestMethod1";
            
            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[1].Tags.First().Name, Is.EqualTo("XYZ"));

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result, "ABC");

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(6));
        }

        [Test]
        public void RunRowTest_runOnlyWithTagABCandXYZ_CheckTagAndCheckValue2()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Tags(""ABC"")
'AccUnit:Row(1)
'AccUnit:Row(2).Tags(""XYZ"")
'AccUnit:Row(3)
public Function TestMethod1(ByVal x as Long) as Long
   m_Check = m_Check + x
   TestMethod1 = x
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var memberName = "TestMethod1";
            
            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[1].Tags.First().Name, Is.EqualTo("XYZ"));

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result, "ABC,XYZ");

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }

        [Test]
        public void TagInClassHeader_runOnlyWithTagABCandXYZ_CheckTagAndCheckValue2()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
'AccUnit:TestClass:Tags(""ABC"")

private m_Check as Long

'AccUnit:Row(1)
'AccUnit:Row(2).Tags(""XYZ"")
'AccUnit:Row(3)
public Function TestMethod1(ByVal x as Long) as Long
   m_Check = m_Check + x
   TestMethod1 = x
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var memberName = "TestMethod1";

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[1].Tags.First().Name, Is.EqualTo("XYZ"));

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result, "ABC,XYZ");

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }
    }
}
