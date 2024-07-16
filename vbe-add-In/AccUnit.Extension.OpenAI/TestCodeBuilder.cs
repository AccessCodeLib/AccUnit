using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenAI;
using OpenAI.Chat;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class TestCodeBuilder
    {
        private readonly IOpenAiService _openAiService;
        private readonly ChatClient _chatClient;

        private bool _disableRowTest = false;
        private string _testProcedureName;
        private string _baseProcedureClassName;
        private string _baseProcedureCode;
        private string _testProcedureParameters;

        public TestCodeBuilder(IOpenAiService openAiService)
        {
            _openAiService = openAiService;
            _chatClient = _openAiService.NewChatClient();
        }

        public TestCodeBuilder DisableRowTest()
        {
            _disableRowTest = true;
            return this;
        }

        public TestCodeBuilder ProcedureToTest(string procedureCode, string className = null)
        {
            _baseProcedureClassName = className;
            _baseProcedureCode = procedureCode;
            return this;
        }

        public TestCodeBuilder TestProcedureName(string testProcedureName)
        {
            _testProcedureName = testProcedureName;
            return this;
        }

        public TestCodeBuilder TestProcedureParameters(string testProcedureParameters)
        {
            _testProcedureParameters = testProcedureParameters;
            return this;
        }

        public string BuildTestProcedureCode()
        {
            var procMessage = string.IsNullOrEmpty(_baseProcedureClassName)
                ? ProcedureTemplate.Replace("{METHODCODE}", _baseProcedureCode)
                : ProcedureTemplateWithClassName.Replace("{METHODCODE}", _baseProcedureCode).Replace("{CLASSNAME}", _baseProcedureClassName);

            var prePrompt = _disableRowTest ? SimpleTestPrePrompt : RowTestPrePrompt;

            var messages = new List<UserChatMessage>
            {
                new UserChatMessage(prePrompt),
                new UserChatMessage(procMessage)
            };

            if (_testProcedureName != null)
            {
                messages.Add(new UserChatMessage(TestProcedureNameTemplate.Replace("{TESTMETHODNAME}", _testProcedureName)));
            }

            if (_testProcedureParameters != null)
            {
                messages.Add(new UserChatMessage(TestProcedureParametersTemplate.Replace("{PARAMETERS}", _testProcedureParameters)));
            }

            ChatCompletion chatCompletion = _chatClient.CompleteChat(messages);
            var testCode = chatCompletion.Content[0].Text;

            return CleanCode(testCode);
        }

        private string CleanCode(string code)
        {
            code = code.Replace("\r\n", "\n");

            if (code.StartsWith("```"))
            {
                code = code.Substring(code.IndexOf("\n") + 1);
            }
            if (code.EndsWith("```"))
            {
                code = code.Substring(0, code.LastIndexOf("\n"));
            }

            return code.Replace("\n", "\r\n");
        }

        const string SimpleTestPrePrompt = @"
I aim to create a test procedure similar to NUnit.
I work with VBA in Access and utilize the AccUnit testing framework.
Please use the following format for the test: 

```vba
Public Sub TestMethod()
    ' Arrange
    ...
    ' Act
    ...
    ' Assert
    Assert.That ...
End Sub
```

Return only the code without explanation.
Note for assert: since Is is not allowed as a variable in VBA, the framework uses Iz (e.g. for Iz.EqualTo) as a substitute. 
Please create a test procedure for the following method.
";

        const string RowTestPrePrompt = @"
I aim to create a test procedure that uses row-test definitions similar to NUnit.
I work with VBA in Access and utilize the AccUnit testing framework.
I expect each AccUnit:Row entry to be treated as a separate test case, and for the test results to be checked directly within the test method itself.
Please use the following format for the test: 

```vba
'AccUnit:Row(<param1>, <param2>, ... , ExpectedValue).Name(...)
'AccUnit:Row(...)
Public Sub TestMethod(...)
    ' Arrange
    ...
    ' Act
    ...
    ' Assert
    Assert.That ...
End Sub
```

Parameters should be directly included in the signature of the test procedure. Also use an Expected parameter and define the value in the test row definition. Set optional parameters to required.
The AccUnit:Row annotations should be defined outside the procedure. 
No AccUnit:Row if method has no parameters.
No blank line between row lines and procedure declaration.
Return only the code without explanation.
Note for assert: since Is is not allowed as a variable in VBA, the framework uses Iz (e.g. for Iz.EqualTo) as a substitute. 
Please create a test procedure for the following method.
";

        const string ProcedureTemplate = @"
Please create a test procedure for the following method: 
{METHODCODE}
";

        const string ProcedureTemplateWithClassName = @"
Please create a test procedure for the following method from the {CLASSNAME} class: 
{METHODCODE}
";

        const string TestProcedureNameTemplate = @"
Use {TESTMETHODNAME} as name for test method.
";
        const string TestProcedureParametersTemplate = @"
Use {PARAMETERS} as parameters for test method.
";

    }
}
