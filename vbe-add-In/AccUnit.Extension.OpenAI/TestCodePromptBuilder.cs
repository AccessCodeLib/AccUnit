using AccessCodeLib.AccUnit.Extension.OpenAI.Properties;
using Microsoft.Extensions.Primitives;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class TestCodePromptBuilder : ITestCodePromptBuilder
    {
        public string BuildPrePrompt(bool useRowTest, string testMethodTemplate)
        {
            var prePrompt = useRowTest ? Settings.Default.RowTestPrePrompt : Settings.Default.SimpleTestPrePrompt;
            prePrompt = prePrompt.Replace("{TESTMETHODTEMPLATE}", testMethodTemplate ?? Settings.Default.DefaultTestMethodTemplate);

            return prePrompt;
        }

        public string BuildPrompt(string baseProcedureCode, string baseProcedureClassName,
                                  string testMethodName, string testMethodParameters)
        {
            var procMessage = string.IsNullOrEmpty(baseProcedureClassName)
                ? Settings.Default.ProcedureTemplate.Replace("{METHODCODE}", baseProcedureCode)
                : Settings.Default.ProcedureTemplateWithClassName.Replace("{METHODCODE}", baseProcedureCode).Replace("{CLASSNAME}", baseProcedureClassName);


            var sb = new StringBuilder();
            sb.Append(procMessage);
            if (!string.IsNullOrEmpty(testMethodName))
            {
                sb.Append("\n\r" + Settings.Default.TestProcedureNameTemplate.Replace("{TESTMETHODNAME}", testMethodName));
            }
            if (!string.IsNullOrEmpty(testMethodParameters))
            {
                sb.Append("\n\r" + Settings.Default.TestProcedureParametersTemplate.Replace("{PARAMETERS}", testMethodParameters));
            }
            var prompt = sb.ToString();

            return prompt;
        }
    }
}
