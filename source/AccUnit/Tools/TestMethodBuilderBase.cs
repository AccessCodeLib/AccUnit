using System.Linq;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.Tools
{
    public abstract class TestMethodBuilderBase : ITestMethodBuilder
    {
        public abstract string GenerateProcedureCode(TestCodeModuleMember member);

        protected virtual string GetProcedureNameForTest(TestCodeModuleMember member)
        {
            // use Get, Let or Set prefix related to the member type
            var suffix = member.ProcKind == Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Get ? "_Get"
                            : member.ProcKind == Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Let ? "_Let"
                            : member.ProcKind == Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Set ? "_Set" : "";

            return member.Name + suffix;
        }

        protected virtual string GetProcedureParameterString(string procedureName, string procDeclaration)
        {
            if (string.IsNullOrEmpty(procDeclaration))
                return "()";

            var declarationCheckString = procDeclaration.Replace(" ", "");
            if (declarationCheckString.Contains("()") && (declarationCheckString.Count(c => c == '(') == 1))
                return "()";

            procDeclaration = procDeclaration.Replace("Optional ", "").Replace("ParamArray", "ByRef");

            //remove string between "=" and ("," or ")")
            var equalSignIndex = procDeclaration.IndexOf("=");
            while (equalSignIndex > 0)
            {
                var commaIndex = procDeclaration.IndexOf(",", equalSignIndex);

                if (commaIndex > 0)
                {
                    procDeclaration = procDeclaration.Remove(equalSignIndex, commaIndex - equalSignIndex);
                }
                else
                {
                    // issue #7-2: Public Function Xyz(Byval X as long, ParamArray P() As Variant) as String
                    procDeclaration = procDeclaration.Replace("()", "[]");
                    var bracketIndex = procDeclaration.IndexOf(")", equalSignIndex);
                    if (bracketIndex > 0)
                    {
                        procDeclaration = procDeclaration.Remove(equalSignIndex, bracketIndex - equalSignIndex);
                    }
                    procDeclaration = procDeclaration.Replace("[]", "()");
                }
                equalSignIndex = procDeclaration.IndexOf("=");
            }

            var parameters = procDeclaration.Substring(procDeclaration.IndexOf(procedureName) + procedureName.Length);
            parameters = ConvertReturnValueToExpectedWithParam(parameters);
            return parameters;
        }

        private static readonly Regex ConvertReturnValueToExpectedWithParamRegex = new Regex(@"\) As ([^\s]*)", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        protected virtual string ConvertReturnValueToExpectedWithParam(string parameters)
        {
            parameters = parameters.Replace("()", "[]");
            parameters = ConvertReturnValueToExpectedWithParamRegex.Replace(parameters,
                                                               m =>
                                                               string.Format(", ByVal Expected As {0})", m.Groups[1].Value));
            return parameters.Replace("[]", "()");
        }

    }
}
