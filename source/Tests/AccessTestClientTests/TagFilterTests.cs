using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using System.Collections.Generic;
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
        public void GenerateEnumarableTagsFormString_CheckTags()
        {
            object tagString = "abc";
            var tags = Interop.TestRunner.GetFilterTagEnumerableFromObject(tagString);
            Assert.That(tags.Count(), Is.EqualTo(1));
            Assert.That(tags.First().Name, Is.EqualTo(tagString));  
        }

        [Test]
        public void RunSimpleTest_WithoutTagTestNotRun_CheckTagAndCheckValue0()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

public Function TestMethod1()
   m_Check = 2
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var memberName = "TestMethod1";

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, memberName, result, "abc");

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
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

        [Test]
        [TestCase("abc", 6)]
        [TestCase("ABC,XYZ", 2)]
        [TestCase("123", 24)]
        [TestCase("123,XYZ", 8)]
        [TestCase("XYZ", 10)]
        public void TagInClassHeaderAndRow_AddFromVBProject_CheckSumValue(string tagFilter, int valueToCheck)
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "modAccUnitTestClass", vbext_ComponentType.vbext_ct_StdModule, @"

public m_Check as Long
");

            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"

'AccUnit:TestClass:Tags(""ABC"")


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

            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass2", vbext_ComponentType.vbext_ct_ClassModule, @"
'AccUnit:TestClass:Tags(""123"")

'AccUnit:Row(7)
'AccUnit:Row(8).Tags(""XYZ"")
'AccUnit:Row(9)
public Function TestMethod1(ByVal x as Long) as Long
   m_Check = m_Check + x
   TestMethod1 = x
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");

            var testSuite = new VBATestSuite();
            testSuite.ActiveVBProject = _testBuilder.ActiveVBProject;
            testSuite.HostApplication = _accessTestHelper.Application;

            var tagFilters = tagFilter.Split(',');
            var tagList = new List<ITestItemTag>();
            foreach (var tag in tagFilters)
            {
                tagList.Add(new TestItemTag(tag));
            }   

            testSuite.AddFromVBProject();
            testSuite.Filter(tagList);
            testSuite.Run();

            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            var invocHelper = new InvocationHelper(fixture);
            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(valueToCheck));
        }
    }
}
