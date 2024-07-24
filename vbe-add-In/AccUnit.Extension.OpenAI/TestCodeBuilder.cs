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
        private readonly ITestCodePromptBuilder _promptBuilder; 

        public TestCodeBuilderFactory(IOpenAiService openAiService, ITestCodePromptBuilder promptBuilder)
        {
            _openAiService = openAiService;
            _promptBuilder = promptBuilder; 
        }

        public ITestCodeBuilder NewTestCodeBuilder()
        {
           return new TestCodeBuilder(_openAiService, _promptBuilder);
        }
    }

    public class TestCodeBuilder : ITestCodeBuilder
    {
        private readonly IOpenAiService _openAiService;
        private readonly ITestCodePromptBuilder _promptBuilder;

        private bool _disableRowTest = false;
        private string _testMethodTemplate = null;
        private string _testMethodName;
        private string _testMethodParameters;

        private string _baseProcedureClassName;
        private string _baseProcedureCode;

        public TestCodeBuilder(IOpenAiService openAiService, ITestCodePromptBuilder promptBuilder)
        {
            _openAiService = openAiService;
            _promptBuilder = promptBuilder; 
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
            var result = BuildTestMethodCodeAsync(); 
            return result;  
        }

        public string BuildTestMethodCodeAsync()
        {
            var prePrompt = _promptBuilder.BuildPrePrompt(!_disableRowTest, _testMethodTemplate);
            var prompt = _promptBuilder.BuildPrompt(_baseProcedureCode, _baseProcedureClassName, _testMethodName, _testMethodParameters);  

            var messages = new[]
            {
                new { role = "assistant", content = prePrompt },
                new { role = "user", content = prompt }
            };

            var result = _openAiService.SendRequest(messages);
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

        
    }
}
