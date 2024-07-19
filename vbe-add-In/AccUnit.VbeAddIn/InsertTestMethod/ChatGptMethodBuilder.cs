using AccessCodeLib.AccUnit.Extension.OpenAI;
using AccessCodeLib.AccUnit.Tools;

namespace AccessCodeLib.AccUnit.VbeAddIn.InsertTestMethod
{
    public class ChatGptMethodBuilder : TestMethodBuilderBase //: TemplateBasedTestMethodBuilder
    {
        private readonly ITestCodeBuilderFactory _testCodeBuilderFactory;

        public ChatGptMethodBuilder(ITestCodeBuilderFactory testCodeBuilderFactory)
        {
            _testCodeBuilderFactory = testCodeBuilderFactory;
        }

        public override string GenerateProcedureCode(TestCodeModuleMember member)
        {

            //var templateSource = base.GenerateProcedureCode(member);
            
            var testCodeBuilder = _testCodeBuilderFactory.NewTestCodeBuilder();
            var templateSource = TestCodeBuilder.DefaultTestMethodTemplate;

            var codeToTest = member.ProcedureCode;
            if (string.IsNullOrEmpty(codeToTest))
                codeToTest = member.DeclarationString;

            testCodeBuilder.ProcedureToTest(codeToTest, member.CodeModuleName)
                           .TestMethodTemplate(templateSource);
                           
            var testCode = testCodeBuilder.BuildTestMethodCode();

            return testCode;

        }
    }
}
