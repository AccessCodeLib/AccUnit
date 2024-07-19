using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests
{
    public class TestCodeBuilderTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void BuildTestCode_SimpleTest_DefineTestProcName()
        {
            var builder = new TestCodeBuilder(new OpenAiService(new CredentialManager()));

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
                Assert.That(testCode.Contains("'AccUnit:Row"), Is.False);
                Assert.That(testCode.Contains("Public Sub GetDate_CheckIfValueReturnedNot0"), Is.True);
            });
        }

        [Test]
        public void BuildTestCode_RowTest_DefineTestProcName()
        {
            var builder = new TestCodeBuilder(new OpenAiService(new CredentialManager()));

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
                Assert.That(testCode.Contains("'AccUnit:Row"), Is.True);
                Assert.That(testCode.Contains("Public Sub Add_2Params_CheckResult"), Is.True);
            });
        }


        [Test]
        public void BuildTestCode_RowTest_DefineTestProcNameAndParams()
        {
            var builder = new TestCodeBuilder(new OpenAiService(new CredentialManager()));

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
                Assert.That(testCode.Contains("'AccUnit:Row"), Is.True);
                Assert.That(testCode.Contains("Public Sub Add_2Params_CheckResult(ByVal intA As Integer, ByVal intB As Integer, ByVal Expected"), Is.True);
            });
        }
    }
}
