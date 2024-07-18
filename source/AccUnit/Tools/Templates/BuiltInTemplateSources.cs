using AccessCodeLib.Common.Tools.Logging;

namespace AccessCodeLib.AccUnit.Tools.Templates
{
    internal static class BuiltInTemplateSources
    {
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
        
        internal static string SimpleTestClassSource  = TestTemplateSources.TestClassHeader +
                                                                "\r\n" +
                                                                TestCodeGenerator.GenerateTemplateProcedureCode(
                                                                    new TestCodeModuleMember("MethodUnderTest"));
        
        // Test class with RowTest
        internal const string RowTestClassName = "zzzRowTest_RENAMEME";
        internal const string RowTestClassCaption = @"Test Class with RowTest";
        internal static readonly string RowTestClassSource = TestTemplateSources.TestClassHeader + @"
' AccUnit:Tags(Example, Row Test)
' AccUnit:Row(1, 2, 3)
' AccUnit:Row(100, 200, 301).Name = ""Failing test case""
Public Sub MethodUnderTest_StateUnderTest_ExpectedBehaviour(ByVal Number1 As Integer, _
                                                            ByVal Number2 As Integer, _
                                                            ByVal ExpectedResult As Integer)
   ' Arrange
   Dim ActualResult As Integer
   
   ' Act
   ActualResult = Number1 + Number2 ' Call production code here
   
   ' Assert
   Assert.That ActualResult, Iz.EqualTo(ExpectedResult)
End Sub
";
    }
}