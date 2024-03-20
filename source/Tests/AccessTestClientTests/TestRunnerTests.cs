using AccessCodeLib.Common.TestHelpers.AccessRelated;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System.Linq;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class TestRunnerTests
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
        public void CreateTestFromExistingFactoryMethode()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public Function TestMethod() as Long
   TestMethod = 123      
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            Assert.That(fixture, Is.Not.Null);

            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod");

        }

        [Test]
        public void FindAndRunTestMethodesFromTestClass()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
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
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "*", result);

            var actual = result.Results.Count();

            Assert.That(actual, Is.EqualTo(2));

        }

        [Test]
        public void FindAndRunTestMethodesFromTestClassWithSetupAndTeardown()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
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
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "*", result);

            var resultCount = result.Results.Count();

            Assert.That(resultCount, Is.EqualTo(2));

            foreach (var testResult in result.Results)
            {
                var res = testResult as Integration.TestResult;
                Assert.That(res.IsPassed, Is.EqualTo(true), res.Message);
            }


            var invocHelper = new InvocationHelper(fixture);
            var ValueAfterTeardowns = invocHelper.InvokeMethod("TestMethod1");
            Assert.That(ValueAfterTeardowns, Is.EqualTo(1));

        }



        [Test]
        public void FindTestMethodesFromTestClassWithoutTlbInf32()
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

            // get type name from object
            var name = Microsoft.VisualBasic.Information.TypeName(fixture);
            Assert.That(name, Is.EqualTo("clsAccUnitTestClass"));

            var vbc = ((VBProject)_testBuilder.ActiveVBProject).VBComponents.Item(name);
            var codeReader = new CodeModuleReader(vbc.CodeModule);

            var publicMembers = codeReader.Members.FindAll(true).FindAll(m => m.ProcKind == vbext_ProcKind.vbext_pk_Proc);
            Assert.That(publicMembers.Count, Is.EqualTo(2));

            foreach (var member in publicMembers)
            {
                Assert.That(member.Name.Substring(0, 10), Is.EqualTo("TestMethod"));
            }

        }

        [Test]
        public void RunRowTest_1Param()
        {
            var classCodeModule = AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(123).Name = ""Row1""
public Function TestMethod1(byval x as Long) as Long
   m_Check = x   
   TestMethod1 = x
End Function

'AccUnit:Row(123, 2).Name = ""Row1""
public Function TestMethod2(byval x as Long, byval y as Long) as Long
   m_Check = x + y 
   TestMethod2 = x + y
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            Assert.That(fixture, Is.Not.Null);

            var memberName = "TestMethod1";
            var fixtureMember = new TestFixtureMember(memberName);

            var testClassReader = new TestClassReader((VBProject)_testBuilder.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = (VBProject)_testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows.Count, Is.EqualTo(1));
            Assert.That(testRows[0].Args[0], Is.EqualTo(123));

            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(123));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(123));

            testRunner.Run(fixture, "TestMethod2", result);
            var valueAfterTestRun2 = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun2, Is.EqualTo(125));
        }

        [Test]
        public void RunRowTest_ArrayParam()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(New Integer() {1, 2})
public Function TestMethod1(ByRef x() as Long) as Long
   m_Check = x(1) 
   TestMethod1 = x(0)
End Function

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);
            Assert.That(fixture, Is.Not.Null);

            var memberName = "TestMethod1";
            var fixtureMember = new TestFixtureMember(memberName);

            var testClassReader = new TestClassReader((VBProject)_testBuilder.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = (VBProject)_testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(1));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }

        [Test]
        public void RunRowTest_WithVbaConstant()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(vbsunday)
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
            Assert.That(fixture, Is.Not.Null);

            var memberName = "TestMethod1";
            var fixtureMember = new TestFixtureMember(memberName);

            var testClassReader = new TestClassReader((VBProject)_testBuilder.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = (VBProject)_testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);
            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(1));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(1));
        }

        [Test]
        public void RunRowTest_IgnoreRow_CheckValueIs0()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(1).Ignore(""test"")
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
            Assert.That(fixture, Is.Not.Null);

            var memberName = "TestMethod1";
            var fixtureMember = new TestFixtureMember(memberName);

            var testClassReader = new TestClassReader((VBProject)_testBuilder.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            Assert.That(fixtureMember.TestClassMemberInfo.TestRows[0].IgnoreInfo.Ignore, Is.True);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = (VBProject)_testBuilder.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[0].IgnoreInfo.Ignore, Is.True);

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(0));
        }

        [Test]
        public void RunTestMethod_OneMethodOnly_CheckValue()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

public Sub TestMethod1()
   m_Check = m_Check + 1
End Sub

public Sub TestMethod2()
   m_Check = m_Check + 2
End Sub

public Function GetCheckValue() as long
   GetCheckValue = m_Check
End Function
");
            var fixtureName = "clsAccUnitTestClass";
            var fixture = _testBuilder.CreateTest(fixtureName);

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod2", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }

        [Test]
        public void RunTestMethods_2MethodsAsEnumerable_CheckValue()
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
            var fixtureName = "clsAccUnitTestClass";
            var testFixtureInstance = _testBuilder.CreateTest(fixtureName);
            var testFixture = new AccessCodeLib.AccUnit.TestFixture(testFixtureInstance);
            testFixture.FillInstanceMembers((VBProject)_testBuilder.ActiveVBProject);
            testFixture.FillTestListFromTestClassInstance((VBProject)_testBuilder.ActiveVBProject);

            var invocHelper = new InvocationHelper(testFixtureInstance);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            var methods = new string[] { "TestMethod1", "TestMethod2" };
            testRunner.Run(testFixture, result, methods, null);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(3));
        }

        [Test]
        public void RunTestMethods_2MethodsAsEnumerableWithPlaceholder_CheckValue()
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
            var fixtureName = "clsAccUnitTestClass";
            var testFixtureInstance = _testBuilder.CreateTest(fixtureName);
            var testFixture = new AccessCodeLib.AccUnit.TestFixture(testFixtureInstance);
            testFixture.FillInstanceMembers((VBProject)_testBuilder.ActiveVBProject);
            testFixture.FillTestListFromTestClassInstance((VBProject)_testBuilder.ActiveVBProject);

            var invocHelper = new InvocationHelper(testFixtureInstance);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner((VBProject)_testBuilder.ActiveVBProject);
            var methods = new string[] { "Te?t*[13]" };
            testRunner.Run(testFixture, result, methods, null);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(5));
        }
    }
}
