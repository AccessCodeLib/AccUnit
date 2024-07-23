using AccessCodeLib.AccUnit.Tools;
using AccessCodeLib.Common.VBIDETools;
using System;
using NUnit.Framework;
using AccessCodeLib.AccUnit.Extension.OpenAI;
using AccessCodeLib.AccUnit.Extension.OpenAI.Tests.TestSupport;
using AccessCodeLib.AccUnit.VbeAddIn.InsertTestMethod;

namespace AccessCodeLib.Common.OpenAI.Tests
{
    public class ChatGptTestMethodBuilderTests
    {
        private const string TestMethodTemplate = @"Public Sub {MethodUnderTest}_{StateUnderTest}_{ExpectedBehaviour}({Params})
	' Arrange
	Dim Actual As Variant
	' Act
	Actual = ""actual value""
	' Assert
	Assert.That Actual, Iz.EqualTo(Expected)
End Sub
";
       
        [Test]
        public void BuildTestCode_RowTest_DefineTestProcName()
        {
            var codebuilderFactory = new TestCodeBuilderFactory(new OpenAiService(new CredentialManager(), new OpenAiRestApiService()));
            var methodBuilder = new ChatGptMethodBuilder(codebuilderFactory, TestMethodTemplate);
            var procedureCode = @"Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer
    Add = A + B
End Function";
            
            var codeModuleMember = new CodeModuleMember(
                            name: "Add", 
                            procKind: Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc, 
                            isPublic: true, 
                            declarationString: "Public Function Add(ByVal A As Integer, ByVal B As Integer) As Integer", 
                            codeModuleName: "TestClass", 
                            procedureCode: procedureCode);
            
            var testCodeModulMember = new TestCodeModuleMember(codeModuleMember, "2Params", "CheckResult");

            var testCode = methodBuilder.GenerateProcedureCode(testCodeModulMember);   
            Console.WriteLine(testCode);

            Assert.Multiple(() =>
            {
                Assert.That(testCode, Is.Not.Null);
                Assert.That(testCode.Contains("'AccUnit:Row"), Is.True);
                Assert.That(testCode.Contains("Public Sub Add_2Params_CheckResult"), Is.True);
            });
        }
    }
}
