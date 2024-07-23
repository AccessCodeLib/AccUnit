using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public interface ITestCodeBuilderFactory
    {
        ITestCodeBuilder NewTestCodeBuilder();
    }

    public class TestCodeBuilderFactory : ITestCodeBuilderFactory
    {
        private readonly IOpenAiService _openAiService;

        public TestCodeBuilderFactory(IOpenAiService openAiService)
        {
            _openAiService = openAiService;
        }

        public ITestCodeBuilder NewTestCodeBuilder()
        {
           return new TestCodeBuilder(_openAiService);
        }
    }


    public class TestCodeBuilder : ITestCodeBuilder
    {
        private readonly IOpenAiService _openAiService;
  
        private bool _disableRowTest = false;
        private string _testMethodTemplate = null;
        private string _testMethodName;
        private string _testMethodParameters;

        private string _baseProcedureClassName;
        private string _baseProcedureCode;

        public TestCodeBuilder(IOpenAiService openAiService)
        {
            _openAiService = openAiService;
        }

        public ITestCodeBuilder DisableRowTest()
        {
            _disableRowTest = true;
            return this;
        }

        public ITestCodeBuilder ProcedureToTest(string procedureCode, string className = null)
        {
            _baseProcedureClassName = className;
            _baseProcedureCode = procedureCode;
            return this;
        }

        public ITestCodeBuilder TestMethodTemplate(string testMethodTemplate)
        {
            _testMethodTemplate = testMethodTemplate;
            return this;
        }

        public ITestCodeBuilder TestMethodName(string testMethodName)
        {
            _testMethodName = testMethodName;
            return this;
        }

        public ITestCodeBuilder TestMethodParameters(string parameterDefinition)
        {
            _testMethodParameters = parameterDefinition;
            return this;
        }

        public string BuildTestMethodCode()
        {
            var result = BuildTestMethodCodeAsync().Result; 
            return result;  
        }

        public async Task<string> BuildTestMethodCodeAsync()
        {


            var procMessage = string.IsNullOrEmpty(_baseProcedureClassName)
                ? ProcedureTemplate.Replace("{METHODCODE}", _baseProcedureCode)
                : ProcedureTemplateWithClassName.Replace("{METHODCODE}", _baseProcedureCode).Replace("{CLASSNAME}", _baseProcedureClassName);

            var prePrompt = _disableRowTest ? SimpleTestPrePrompt : RowTestPrePrompt;
            prePrompt = prePrompt.Replace("{TESTMETHODTEMPLATE}", _testMethodTemplate ?? DefaultTestMethodTemplate);


            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(_testMethodName))
            {
                sb.Append(TestProcedureNameTemplate.Replace("{TESTMETHODNAME}", _testMethodName));
            }
            if (!string.IsNullOrEmpty(_testMethodParameters))
            {
                sb.Append(TestProcedureParametersTemplate.Replace("{PARAMETERS}", _testMethodParameters));
            }
            sb.Append(procMessage);
            var prompt = sb.ToString();

            var messages = new[]
            {
                new { role = "assistant", content = prePrompt },
                new { role = "user", content = prompt }
            };

            var result = await _openAiService.SendRequest(messages);
            var testCode = result;

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
{TESTMETHODTEMPLATE}
```
"
+ PrePromptEndStatement;

        const string RowTestPrePrompt = @"
I aim to create a test procedure that uses row-test definitions similar to NUnit.
I work with VBA in Access and utilize the AccUnit testing framework.
I expect each AccUnit:Row entry to be treated as a separate test case, and for the test results to be checked directly within the test method itself.
Please use the following format for the test: 

```vba
'AccUnit:Row(<param1>, <param2>, ... , ExpectedValue).Name(...)
'AccUnit:Row(...)
{TESTMETHODTEMPLATE}
```

Parameters should be directly included in the signature of the test procedure. Also use an Expected parameter and define the value in the test row definition. Set optional parameters to required.
Test methods must be declared as Public.
The AccUnit:Row annotations should be defined outside the procedure. 
No AccUnit:Row if method has no parameters.
No blank line between row lines and procedure declaration." 
+ PrePromptEndStatement;

        private const string PrePromptEndStatement = @"
Return only the code without explanation.
Note for assert: since Is is not allowed as a variable in VBA, the framework uses Iz (e.g. for Iz.EqualTo) as a substitute. Don't use Call Assert.That(...). Use only Assert.That ...
Please create a test procedure for the following method.
";

        public const string DefaultTestMethodTemplate = @"
Public Sub TestMethod(...)
    ' Arrange
    ...
    ' Act
    ...
    ' Assert
    Assert.That ...
End Sub
";

        const string ProcedureTemplate = @"
Please create a test procedure for the following method: 
{METHODCODE}
";

        const string ProcedureTemplateWithClassName = @"
Please create a test procedure for the following method from the class {CLASSNAME}: 
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
