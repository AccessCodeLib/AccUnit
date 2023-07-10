using AccessCodeLib.Common.Tools.Logging;

namespace AccessCodeLib.AccUnit.Tools.Templates
{
    static class BuiltInTemplateSources {
        static BuiltInTemplateSources()
        {
            using (new BlockLogger())
            {
                // Just for debugging purposes (timing!)
            }
        }

        // Simple test class
        internal const string SimpleTestClassName = @"zzzSimpleTest_RENAMEME";
        internal const string SimpleTestClassCaption = @"Simple Test Class";
        internal static readonly string SimpleTestClassSource = TestTemplateSources.TestClassHeader +
                                                                "\r\n" +
                                                                TestCodeGenerator.GenerateProcedureCode(
                                                                    new TestCodeModuleMember("MethodUnderTest"));

        // Test class with AccUnit-Features
        internal const string AccUnitTestClassName = "zzzAccUnitTest_RENAMEME";
        internal const string AccUnitTestClassCaption = @"Test Class with AccUnit-Features";
        internal static readonly string AccUnitTestClassSource = TestTemplateSources.TestClassHeaderWithInterfaces +
                                                                 "\r\n' AccUnit:Tags(Example, Simple Test)\r\n" +
                                                                 TestCodeGenerator.GenerateProcedureCode(
                                                                     new TestCodeModuleMember("MethodUnderTest1")) +
                                                                 "\r\n" +
                                                                 "\r\n' AccUnit:Tags(Example, Ignored Test)\r\n" +
                                                                 "' AccUnit:Ignore \"I cannot get this test to work now (but is it really necessary?).\"\r\n" +
                                                                 TestCodeGenerator.GenerateProcedureCode(
                                                                     new TestCodeModuleMember("MethodUnderTest2"));

        // Test class with RowTest
        internal const string RowTestClassName = "zzzRowTest_RENAMEME";
        internal const string RowTestClassCaption = @"Test Class with RowTest";
        internal static readonly string RowTestClassSource = TestTemplateSources.TestClassHeaderWithInterfaces + @"
' AccUnit:Tags(Example, Row Test)
' AccUnit:Row(1, 2, 3)
' AccUnit:Row(100, 200, 301).Name = ""Failing test case""
Public Sub MethodUnderTest_StateUnderTest_ExpectedBehaviour(ByVal number1 As Integer, _
                                                            ByVal number2 As Integer, _
                                                            ByVal expectedResult As Integer)
   ' Arrange
   Dim actualResult As Integer
   
   ' Act
   actualResult = number1 + number2 ' Call production code here
   
   ' Assert
   Assert.That actualResult, Iz.EqualTo(expectedResult)
End Sub
";
    }
}