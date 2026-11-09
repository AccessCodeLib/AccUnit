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
            var applicationHelper = new AccessApplicationHelper(_accessTestHelper.Application);
            _testBuilder = new Interop.TestBuilder(applicationHelper);
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

            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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

            var vbc = _accessTestHelper.ActiveVBProject.VBComponents.Item(name);
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

            var testClassReader = new TestClassReader(_accessTestHelper.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _accessTestHelper.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows.Count, Is.EqualTo(1));
            Assert.That(testRows[0].Args[0], Is.EqualTo(123));

            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(123));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(123));

            testRunner.Run(fixture, "TestMethod2", result);
            var valueAfterTestRun2 = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun2, Is.EqualTo(125));
        }

        [Test]
        public void RunTests_SetupRaisesError_TestIsErrorAndTeardownRunsAndRunContinues()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_TeardownCount as Long
private m_TestCount as Long

public sub Setup()
   Err.Raise 5, , ""setup failed""
End sub

public sub Teardown()
   m_TeardownCount = m_TeardownCount + 1
End sub

public Sub TestMethod1()
   m_TestCount = m_TestCount + 1
End Sub

public Sub TestMethod2()
   m_TestCount = m_TestCount + 1
End Sub

public Function GetTeardownCount() as Long
   GetTeardownCount = m_TeardownCount
End Function

public Function GetTestCount() as Long
   GetTestCount = m_TestCount
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "*", result);

            // public helper functions (GetTeardownCount, GetTestCount) are run as tests, too
            Assert.That(result.Results.Count(), Is.EqualTo(4));
            foreach (var testResult in result.Results)
            {
                Assert.That(testResult.IsError, Is.True);
                Assert.That(testResult.IsPassed, Is.False);
                Assert.That(testResult.Message, Does.StartWith("SetUp:"));
            }

            var invocHelper = new InvocationHelper(fixture);
            Assert.That(invocHelper.InvokeMethod("GetTeardownCount"), Is.EqualTo(4));
            Assert.That(invocHelper.InvokeMethod("GetTestCount"), Is.EqualTo(0));
        }

        [Test]
        public void RunSingleTestMethod_CallsSetupAndTeardown()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Value as Long
private m_TeardownCount as Long

public sub Setup()
   m_Value = 123
End sub

public sub Teardown()
   m_TeardownCount = m_TeardownCount + 1
End sub

public Function TestMethod1() as Long
   TestMethod1 = m_Value
End Function

public Function GetTeardownCount() as Long
   GetTeardownCount = m_TeardownCount
End Function
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            var testResult = testRunner.Run(fixture, "TestMethod1");

            Assert.That(testResult.IsPassed, Is.True, testResult.Message);

            var invocHelper = new InvocationHelper(fixture);
            Assert.That(invocHelper.InvokeMethod("GetTeardownCount"), Is.EqualTo(1));
            // Setup has set m_Value, so a direct call after teardown still shows it
            Assert.That(invocHelper.InvokeMethod("TestMethod1"), Is.EqualTo(123));
        }

        [Test]
        public void RunTest_TeardownRaisesError_PassedTestBecomesError()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public sub Teardown()
   Err.Raise 5, , ""teardown failed""
End sub

public Sub TestMethod1()
End Sub
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            var collector = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "*", collector);
            var testResult = collector.Results.Single();

            Assert.That(testResult.IsError, Is.True);
            Assert.That(testResult.IsPassed, Is.False);
            Assert.That(testResult.Message, Does.Contain("TearDown:"));
        }

        [Test]
        public void RunTest_TeardownRaisesError_FailureIsKeptAndMessageAppended()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public sub Teardown()
   Err.Raise 5, , ""teardown failed""
End sub

public Sub TestMethod1()
   Err.Raise 5, , ""test failed""
End Sub
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            var collector = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "*", collector);
            var testResult = collector.Results.Single();

            Assert.That(testResult.IsPassed, Is.False);
            Assert.That(testResult.IsError || testResult.IsFailure, Is.True);
            Assert.That(testResult.Message, Does.Contain("test failed"));
            Assert.That(testResult.Message, Does.Contain("TearDown:"));
        }

        [Test]
        public void RunRowTest_SetupRaisesError_EachRowGetsOwnErrorResult()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
public sub Setup()
   Err.Raise 5, , ""setup failed""
End sub

'AccUnit:Row(1)
'AccUnit:Row(2)
public Sub TestMethod1(byval x as Long)
End Sub
");
            var fixture = _testBuilder.CreateTest("clsAccUnitTestClass");
            var collector = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "*", collector);
            var rowResult = (Integration.TestResultCollection)collector.Results.Single();

            Assert.That(rowResult.Count, Is.EqualTo(2));
            Assert.That(rowResult.All(r => r.IsError), Is.True);
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

            var testClassReader = new TestClassReader(_accessTestHelper.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _accessTestHelper.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(1));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            testRunner.Run(fixture, "TestMethod1", result);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(2));
        }

        [Test]
        public void RunRowTest_VbaArrayParam()
        {
            AccessClientTestHelper.CreateTestCodeModule(_accessTestHelper, "clsAccUnitTestClass", vbext_ComponentType.vbext_ct_ClassModule, @"
private m_Check as Long

'AccUnit:Row(Array(1, 2))
public Function TestMethod1(ByRef x() as Variant) as Long
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

            var testClassReader = new TestClassReader(_accessTestHelper.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _accessTestHelper.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(1));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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

            var testClassReader = new TestClassReader(_accessTestHelper.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _accessTestHelper.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);
            var invocHelper = new InvocationHelper(fixture);
            var returnValue = invocHelper.InvokeMethod("TestMethod1", testRows[0].Args.ToArray());
            Assert.That(returnValue, Is.EqualTo(1));

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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

            var testClassReader = new TestClassReader(_accessTestHelper.ActiveVBProject);
            fixtureMember.TestClassMemberInfo = testClassReader.GetTestClassMemberInfo(fixtureName, memberName);

            Assert.That(fixtureMember.TestClassMemberInfo.TestRows[0].IgnoreInfo.Ignore, Is.True);

            var rowGenerator = new TestRowGenerator
            {
                ActiveVBProject = _accessTestHelper.ActiveVBProject,
                TestName = fixtureName
            };
            var testRows = rowGenerator.GetTestRows(memberName);

            Assert.That(testRows[0].IgnoreInfo.Ignore, Is.True);

            var invocHelper = new InvocationHelper(fixture);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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
            testFixture.FillInstanceMembers(_accessTestHelper.ActiveVBProject);
            testFixture.FillTestListFromTestClassInstance(_accessTestHelper.ActiveVBProject);

            var invocHelper = new InvocationHelper(testFixtureInstance);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
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
            testFixture.FillInstanceMembers(_accessTestHelper.ActiveVBProject);
            testFixture.FillTestListFromTestClassInstance(_accessTestHelper.ActiveVBProject);

            var invocHelper = new InvocationHelper(testFixtureInstance);

            var result = new TestResultCollector();
            var testRunner = new Interop.TestRunner(_accessTestHelper.ActiveVBProject);
            var methods = new string[] { "Te?t*[13]" };
            testRunner.Run(testFixture, result, methods, null);

            var valueAfterTestRun = invocHelper.InvokeMethod("GetCheckValue");
            Assert.That(valueAfterTestRun, Is.EqualTo(5));
        }
    }
}
