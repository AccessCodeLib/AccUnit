using AccessCodeLib.AccUnit.Tools;
using AccessCodeLib.Common.VBIDETools;
using NUnit.Framework;

namespace TestGeneratorTests
{
    public class TestCodeGeneratorTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void ConvertReturnValueToExpectedWithParam_WithSimpleParam_ReturnWithExpected()
        {
            var baseString = "Public Function Xyz(ByVal x As Long) As String";
            var expected = "Public Function Xyz(ByVal x As Long, ByVal Expected As String)";

            var actual = TestCodeGenerator.ConvertReturnValueToExpectedWithParam(baseString);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConvertReturnValueToExpectedWithParam_WithArray_ReturnWithExpected()
        {
            var baseString = "Public Function StringFormat2(ByVal s As String, ByRef x() As Variant) As String";
            var expected = "Public Function StringFormat2(ByVal s As String, ByRef x() As Variant, ByVal Expected As String)";

            var actual = TestCodeGenerator.ConvertReturnValueToExpectedWithParam(baseString);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetProcedureParameterString_WithArray_ReturnWithExpected()
        {
            const string procName = "StringFormat2";
            const string baseString = "Public Function " + procName + "(ByVal s As String, ByRef x() As Variant) As String";
            var expected = "(ByVal s As String, ByRef x() As Variant, ByVal Expected As String)";

            var actual = TestCodeGenerator.GetProcedureParameterString(procName, baseString);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetProcedureParameterString_WithParamArray_ReturnWithByRefArray()
        {
            const string procName = "StringFormat2";
            const string baseString = "Public Function " + procName + "(ByVal s As String, ParamArray x() As Variant) As String";
            var expected = "(ByVal s As String, ByRef x() As Variant, ByVal Expected As String)";

            var actual = TestCodeGenerator.GetProcedureParameterString(procName, baseString);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetProcedureRowTestString_WithParamArray_ReturnWithArray()
        {
            const string parameters = "(ByVal s As String, ByRef x() As Variant, ByVal Expected As String)";
            var expected = "'AccUnit:Row(s, x(), Expected).Name = \"Example row - please replace the parameter names with values)\"";

            var actual = TestCodeGenerator.GetProcedureRowTestString(parameters);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetProcedureRowTestString_WithParamArrayAndComment_ReturnWithArray()
        {
            const string parameters = "(ByVal s As String, ByRef x() As Variant, ByVal Expected As String) ' this is a commend";
            var expected = "'AccUnit:Row(s, x(), Expected).Name = \"Example row - please replace the parameter names with values)\"";

            var actual = TestCodeGenerator.GetProcedureRowTestString(parameters);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateProcedureCode_WithoutParam_ReturnExpectedDeclarationInBody()
        {
            var codeModuleMember = new CodeModuleMember("Method1", Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc, true, "Public Function Method1() As String");
            var expected = @"Public Sub Method1()
" + "\t" + @"' Arrange
" + "\t" + @"Err.Raise vbObjectError, ""Method1"", ""Test not implemented""
" + "\t" + @"Const Expected As Variant = ""expected value""
" + "\t" + @"Dim Actual As Variant
" + "\t" + @"' Act
" + "\t" + @"Actual = ""actual value""
" + "\t" + @"' Assert
" + "\t" + @"Assert.That Actual, Iz.EqualTo(Expected)
End Sub
";

            var actual = TestCodeGenerator.GenerateProcedureCode(codeModuleMember, "", "");

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateProcedureCode_WithParam_ReturnWithoutExpectedDeclarationInBody()
        {
            var codeModuleMember = new CodeModuleMember("Method1", Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc, true, "Public Function Method1(ByVal x As Long) As String");
            var expected = @"'AccUnit:Row(x, Expected).Name = ""Example row - please replace the parameter names with values)""
Public Sub Method1(ByVal x As Long, ByVal Expected As String)
" + "\t" + @"' Arrange
" + "\t" + @"Err.Raise vbObjectError, ""Method1"", ""Test not implemented""
" + "\t" + @"Dim Actual As Variant
" + "\t" + @"' Act
" + "\t" + @"Actual = ""actual value""
" + "\t" + @"' Assert
" + "\t" + @"Assert.That Actual, Iz.EqualTo(Expected)
End Sub
";

            var actual = TestCodeGenerator.GenerateProcedureCode(codeModuleMember, "", "");

            Assert.AreEqual(expected, actual);
        }


    }
}