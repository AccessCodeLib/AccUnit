using AccessCodeLib.AccUnit.Extension.OpenAI;
using AccessCodeLib.AccUnit.Tools;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn.InsertTestMethod
{
    public class ChatGptMethodBuilder : TemplateBasedTestMethodBuilder
    {
        private readonly ITestCodeBuilderFactory _testCodeBuilderFactory;

        public ChatGptMethodBuilder(ITestCodeBuilderFactory testCodeBuilderFactory, string testMethodTemplate)
            : base(testMethodTemplate)
        {
            _testCodeBuilderFactory = testCodeBuilderFactory;
        }

        public override string GenerateProcedureCode(TestCodeModuleMember member)
        {
            var templateSource = base.GenerateProcedureCode(member);
            var testMethodName = GetTestMethodNameFromSource(templateSource);

            var codeToTest = member.ProcedureCode;
            if (string.IsNullOrEmpty(codeToTest))
                codeToTest = member.DeclarationString;

            string testCode;

            //UITools.ShowMessage(templateSource);

            try
            {
                var testCodeBuilder = _testCodeBuilderFactory.NewTestCodeBuilder();
                testCodeBuilder.ProcedureToTest(codeToTest, member.CodeModuleName)
                           .TestMethodTemplate(templateSource)
                           .TestMethodName(testMethodName);

                //UITools.ShowMessage("now build code ..");
                 testCode = testCodeBuilder.BuildTestMethodCode();
            }
            catch (Exception ex)
            {
                UITools.ShowException(ex);
                testCode = templateSource;
            }

            return testCode;
        }

        private string GetTestMethodNameFromSource(string templateSource)
        {
            var lines = templateSource.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            // declaration = first line which begins with "Public Sub"
            var declaration = lines.FirstOrDefault(l => l.Trim().StartsWith("Public Sub"));

            // method name = word after "Public Sub" and before "("
            var methodName = declaration.PadLeft(declaration.Length - "Public Sub".Length).Split('(')[0].Trim();

            return methodName;
        }

    }
}
