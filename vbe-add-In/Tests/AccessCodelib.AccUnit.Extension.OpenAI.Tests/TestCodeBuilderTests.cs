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
        public void BuildTestCode_SimpleTest_DefineTestProcNameAsClassMember()
        {
            var procedureCode = @"Public Function GetDate() As Date
    GetDate = Date()
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, true, "TestClass")
                                  .TestMethodName("GetDate_CheckIfValueReturnedNot0")
                                  .DisableRowTest()
                                  .BuildTestMethodCodeAsync().Result;
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Not.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub GetDate_CheckIfValueReturnedNot0"));
                Assert.That(testCode, Does.Contain("New TestClass"));
            });
        }

        [Test]
        public void BuildTestCode_SimpleTest_DefineTestProcName()
        {
            var procedureCode = @"Public Function GetDate() As Date
    GetDate = Date()
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, false, "TestModule")
                                  .TestMethodName("GetDate_CheckIfValueReturnedNot0")
                                  .DisableRowTest()
                                  .BuildTestMethodCode();
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode, Does.Not.Contain("'AccUnit:Row"));
                Assert.That(testCode, Does.Contain("Public Sub GetDate_CheckIfValueReturnedNot0"));
                Assert.That(testCode, Does.Not.Contain("New TestModule"));
            });
        }

        [Test]
        public void BuildTestCode_RowTest_DefineTestProcName()
        {
            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, true, "TestClass")
                                  .TestMethodName("Add_2Params_CheckResult")
                                  .BuildTestMethodCodeAsync().Result;
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

            var testCode = builder.ProcedureToTest(procedureCode, true, "TestClass")
                                  .TestMethodName("Add_2Params_CheckResult")
                                  .TestMethodParameters("ByVal intA As Integer, ByVal intB As Integer")
                                  .BuildTestMethodCodeAsync().Result;
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode.Contains("'AccUnit:Row"), Is.True);
                Assert.That(testCode.Contains("Public Sub Add_2Params_CheckResult(ByVal intA As Integer, ByVal intB As Integer, ByVal Expected"), Is.True);
            });
        }

        [Test]
        public void BuildTestCode_SyncRowTest_DefineTestProcNameAndParams()
        {
            var builder = new TestCodeBuilder(new OpenAiService(new CredentialManager(), new OpenAiRestApiService()), new TestCodePromptBuilder());

            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";

            var testCode = builder.ProcedureToTest(procedureCode, true, "TestClass")
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