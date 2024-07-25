using NUnit.Framework;
using System;

namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests
{
    public class TestCodeBuilderTests
    {
        private TestCodeBuilder builder;

        [SetUp]
        public void Setup()
        {
            builder = new TestCodeBuilder(new OpenAiService(new CredentialManager(), new OpenAiRestApiService()), new TestCodePromptBuilder());
        }

        [Test]
        public void BuildTestCode_SimpleTest_DefineTestProcName()
        {
            var procedureCode = @"Public Function GetDate() As Date
    GetDate = Date()
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, "TestClass")
                                  .TestMethodName("GetDate_CheckIfValueReturnedNot0")
                                  .DisableRowTest()
                                  .BuildTestMethodCode();
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Not.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub GetDate_CheckIfValueReturnedNot0"));
            });
        }

        [Test]
        public void BuildTestCode_RowTest_DefineTestProcName()
        {
            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, "TestClass")
                                  .TestMethodName("Add_2Params_CheckResult")
                                  .BuildTestMethodCode();
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub Add_2Params_CheckResult"));
            });
        }


        [Test]
        public void BuildTestCode_RowTest_DefineTestProcNameAndParams()
        {
            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, "TestClass")
                                  .TestMethodName("Add_2Params_CheckResult")
                                  .TestMethodParameters("ByVal intA As Integer, ByVal intB As Integer")
                                  .BuildTestMethodCode();
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub Add_2Params_CheckResult(ByVal intA As Integer, ByVal intB As Integer, ByVal Expected"));
            });
        }

        [Test]
        public void BuildTestCode_SyncRowTest_DefineTestProcNameAndParams()
        {
            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, "TestClass")
                                  .TestMethodName("Add_2Params_CheckResult")
                                  .TestMethodParameters("ByVal intA As Integer, ByVal intB As Integer")
                                  .BuildTestMethodCode();
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub Add_2Params_CheckResult(ByVal intA As Integer, ByVal intB As Integer, ByVal Expected"));
            });
        }
    }
}