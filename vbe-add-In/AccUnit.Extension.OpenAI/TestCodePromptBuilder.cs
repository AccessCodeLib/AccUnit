using System.Text;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class TestCodePromptBuilder : ITestCodePromptBuilder
    {
        public string BuildPrePrompt(bool useRowTest, string testMethodTemplate)
        {
            var prePrompt = useRowTest ? RowTestPrePrompt : SimpleTestPrePrompt;
            prePrompt = prePrompt.Replace("{TESTMETHODTEMPLATE}", testMethodTemplate ?? DefaultTestMethodTemplate);

            return prePrompt;
        }

        public string BuildPrompt(string baseProcedureCode, bool isClassMember, string baseProcedureVbComponentName,
                                  string testMethodName, string testMethodParameters)
        {
            var procMessage = isClassMember
                ? ProcedureTemplateWithClassName.Replace("{METHODCODE}", baseProcedureCode).Replace("{CLASSNAME}", baseProcedureVbComponentName)
                : ProcedureTemplate.Replace("{METHODCODE}", (string.IsNullOrEmpty(baseProcedureVbComponentName) ? "" : baseProcedureVbComponentName + ".") + baseProcedureCode);

            var sb = new StringBuilder();
            sb.Append(procMessage);
            if (!string.IsNullOrEmpty(testMethodName))
            {
                sb.Append("\n\r" + TestProcedureNameTemplate.Replace("{TESTMETHODNAME}", testMethodName));
            }
            if (!string.IsNullOrEmpty(testMethodParameters))
            {
                sb.Append("\n\r" + TestProcedureParametersTemplate.Replace("{PARAMETERS}", testMethodParameters));
            }
            var prompt = sb.ToString();

            return prompt;
        }

        const string SimpleTestPrePrompt = @"Create a test procedure similar to NUnit.
Work with VBA in Access and utilize the AccUnit testing framework.
Please use the following format for the test: 

```vba
{TESTMETHODTEMPLATE}
```
" + PrePromptEndStatement;

        const string RowTestPrePrompt = @"Create a test procedure that uses row-test definitions similar to NUnit.
Work with VBA in Access and utilize the AccUnit testing framework.
I expect each AccUnit:Row entry to be treated as a separate test case, and for the test results to be checked directly within the test method itself.
Please use the following format for the test: 

```vba
'AccUnit:Row([param1], [param2], ... , [ExpectedValue]).Name(...)
'AccUnit:Row([param1], [param2], ... , [ExpectedValue]).Name(...)
{TESTMETHODTEMPLATE}
```

Parameters should be directly included in the signature of the test procedure. Also use an Expected parameter and define the value in the test row definition. Set optional parameters to required.
Test methods must be declared as Public.
The AccUnit:Row annotations should be defined outside the procedure. 
Insert no AccUnit:Row lines if procedure is without parameters.
No blank line between row lines and procedure declaration.
" + PrePromptEndStatement;

        private const string PrePromptEndStatement = @"Return only the code without explanation.
Note for assert: since Is is not allowed as a variable in VBA, the framework uses Iz (e.g. for Iz.EqualTo) as a substitute. Don't use Call Assert.That(...). Use only Assert.That ...";

        private const string DefaultTestMethodTemplate = @"Public Sub TestMethod(...)
    ' Arrange
    ...
    ' Act
    ...
    ' Assert
    Assert.That ...
End Sub";

        private const string ProcedureTemplate = @"Please create a test procedure for the following method: 
{METHODCODE}";

        private const string ProcedureTemplateWithClassName = @"Please create a test procedure with a new class instance for the following method from the class {CLASSNAME}: 
{METHODCODE}";

        private const string TestProcedureNameTemplate = @"Use {TESTMETHODNAME} as name for test method.";
        private const string TestProcedureParametersTemplate = @"Use {PARAMETERS} as parameters for test method.";
    }
}
