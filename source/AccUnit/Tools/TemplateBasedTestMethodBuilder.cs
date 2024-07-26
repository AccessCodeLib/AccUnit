using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using System;

namespace AccessCodeLib.AccUnit.Tools
{
    public class TemplateBasedTestMethodBuilder : TestMethodBuilderBase
    {

        //private static string TestMethodTemplate { get { return TemplatesUserSettings.Current.TestMethodTemplate; } }
        private readonly string _testMethodTemplate;

        public const string TestMethodNameFormat = @"{0}_{1}_{2}"; // methodsUnderTest_stateUnderTest_expectedBehaviour
        // The placeholder constants are public for testing purposes. InternalsVisibleTo seems to be not working here.
        public const string MethodUnderTestPlaceholder = "{MethodUnderTest}";
        public const string StateUnderTestPlaceholder = "{StateUnderTest}";
        public const string ExpectedBehaviourPlaceholder = "{ExpectedBehaviour}";
        public const string ParamsPlaceholder = "({Params})";

        public TemplateBasedTestMethodBuilder(string testMethodTemplate)
        {
            _testMethodTemplate = testMethodTemplate;
        }

        public override string GenerateProcedureCode(TestCodeModuleMember member)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, member.StateUnderTest, member.ExpectedBehaviour)))
            {
                var code = _testMethodTemplate;

                if (string.IsNullOrEmpty(member.StateUnderTest))
                    code = code.Replace("_" + StateUnderTestPlaceholder, StateUnderTestPlaceholder);
                if (string.IsNullOrEmpty(member.ExpectedBehaviour))
                    code = code.Replace("_" + ExpectedBehaviourPlaceholder, ExpectedBehaviourPlaceholder);

                code = code.Replace(MethodUnderTestPlaceholder, GetProcedureNameForTest(member));
                code = code.Replace(StateUnderTestPlaceholder, member.StateUnderTest);
                code = code.Replace(ExpectedBehaviourPlaceholder, member.ExpectedBehaviour);

                var parameters = GetProcedureParameterString(member.Name, member.DeclarationString);
                code = code.Replace(ParamsPlaceholder, parameters);
                if (parameters.Length > 2) // () is the shortest possible parameter string
                {
                    // replace expected declaration in code
                    code = code.Replace("\r\n\tConst Expected As Variant = \"expected value\"", "");

                    // insert row test code
                    code = GetProcedureRowTestString(parameters) + Environment.NewLine + code;
                }

                return code;
            }
        }

        public string GetProcedureRowTestString(string parameters)
        {
            var paramString = parameters.Replace("ByRef ", "").Replace("ByVal ", "").Replace("_" + Environment.NewLine, "")
                    .Replace("()", "[]");

            if (paramString.Contains(")"))
                paramString = paramString.Substring(0, paramString.IndexOf(")"));

            if (paramString.Contains("("))
                paramString = paramString.Substring(1, paramString.Length - 1);


            var Params = paramString.Split(',');

            for (int i = 0; i < Params.Length; i++)
            {
                var param = Params[i];
                if (param.Contains(" As "))
                    param = param.Substring(0, param.IndexOf(" As "));

                Params[i] = param.Trim();
            }
            return @"'AccUnit:Row(" + string.Join(", ", Params).Replace("[]", "()") + ").Name = \"Example row - please replace the parameter names with values)\"";
        }
    }
}
