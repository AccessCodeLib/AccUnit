using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AccessCodeLib.AccUnit.Tools
{
    public interface ITestMethodBuilder
    {
        string GenerateProcedureCode(TestCodeModuleMember member);
    }

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

    public class TemplateBasedTestMethodBuilder : TestMethodBuilderBase
    {
        public const string TestMethodNameFormat = @"{0}_{1}_{2}"; // methodsUnderTest_stateUnderTest_expectedBehaviour
        // The placeholder constants are public for testing purposes. InternalsVisibleTo seems to be not working here.
        public const string MethodUnderTestPlaceholder = "{MethodUnderTest}";
        public const string StateUnderTestPlaceholder = "{StateUnderTest}";
        public const string ExpectedBehaviourPlaceholder = "{ExpectedBehaviour}";
        public const string ParamsPlaceholder = "({Params})";

        public override string GenerateProcedureCode(TestCodeModuleMember member)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, member.StateUnderTest, member.ExpectedBehaviour)))
            {
                var code = TestMethodTemplate;

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

        public static string GetProcedureRowTestString(string parameters)
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

        private static string TestMethodTemplate { get { return TemplatesUserSettings.Current.TestMethodTemplate; } }
    }

    public class TestCodeGenerator
    {
        public const string TestMethodNameFormat = @"{0}_{1}_{2}"; // methodsUnderTest_stateUnderTest_expectedBehaviour
        // The placeholder constants are public for testing purposes. InternalsVisibleTo seems to be not working here.
        public const string MethodUnderTestPlaceholder = "{MethodUnderTest}";
        public const string StateUnderTestPlaceholder = "{StateUnderTest}";
        public const string ExpectedBehaviourPlaceholder = "{ExpectedBehaviour}";
        public const string ParamsPlaceholder = "({Params})";
        private readonly List<TestCodeModuleMember> _codeModuleMembers = new List<TestCodeModuleMember>();

        private readonly ITestMethodBuilder _testMethodBuilder;

        public TestCodeGenerator(ITestMethodBuilder testMethodBuilder)
        {
            _testMethodBuilder = testMethodBuilder;
        }

        public void Add(IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            Add(CreateMembers(methodsUnderTest, stateUnderTest, expectedBehaviour));
        }

        public void Add(IEnumerable<CodeModuleMember> codeModuleMembers)
        {
            Add(codeModuleMembers.Select(member => (member is TestCodeModuleMember testCodeModulMember)
                                                       ? testCodeModulMember
                                                       : new TestCodeModuleMember(member)
                    ));
        }

        private void Add(IEnumerable<TestCodeModuleMember> codeModuleMembers)
        {
            _codeModuleMembers.AddRange(codeModuleMembers);
        }

        private static IEnumerable<CodeModuleMember> CreateMembers(IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            var list = new List<CodeModuleMember>();
            foreach (var method in methodsUnderTest)
            {
                using (new BlockLogger(string.Format(TestMethodNameFormat, method, stateUnderTest, expectedBehaviour)))
                {
                    list.Add(new TestCodeModuleMember(method, stateUnderTest, expectedBehaviour));
                }

            }
            return list;
        }

        public string GenerateSourceCode(bool includeHeader = true)
        {
            var sb = new StringBuilder();
            if (includeHeader)
                sb.AppendLine(TestClassHeader);
            if (_codeModuleMembers != null)
                AddMemberCode(sb);
            return sb.ToString();
        }

        private void AddMemberCode(StringBuilder sb)
        {
            using (new BlockLogger())
            {
                foreach (var codeModuleMemberInfo in _codeModuleMembers)
                {
                    sb.AppendLine();
                    sb.AppendLine(GenerateProcedureCode(codeModuleMemberInfo));
                }
            }
        }

        public string GenerateProcedureCode(CodeModuleMember member,
                                                   string stateUnderTest,
                                                   string expectedBehaviour)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, stateUnderTest, expectedBehaviour)))
            {
                return GenerateProcedureCode(new TestCodeModuleMember(member, stateUnderTest, expectedBehaviour));
            }
        }

        internal string GenerateProcedureCode(TestCodeModuleMember member)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, member.StateUnderTest, member.ExpectedBehaviour)))
            {
                return _testMethodBuilder.GenerateProcedureCode(member);;
            }
        }

        
        internal static string TestClassHeader { get { return TestTemplateSources.TestClassHeader; } }
    }
}
