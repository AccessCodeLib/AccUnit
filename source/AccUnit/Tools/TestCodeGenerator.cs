﻿using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.Tools
{
    public class TestCodeGenerator
    {
        public const string TestMethodNameFormat = @"{0}_{1}_{2}"; // methodsUnderTest_stateUnderTest_expectedBehaviour
        // The placeholder constants are public for testing purposes. InternalsVisibleTo seems to be not working here.
        public const string MethodUnderTestPlaceholder = "{MethodUnderTest}";
        public const string StateUnderTestPlaceholder = "{StateUnderTest}";
        public const string ExpectedBehaviourPlaceholder = "{ExpectedBehaviour}";
        public const string ParamsPlaceholder = "({Params})";
        private readonly List<TestCodeModuleMember> _codeModuleMembers = new List<TestCodeModuleMember>();

        public void Add(IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            Add(CreateMembers(methodsUnderTest, stateUnderTest, expectedBehaviour));
        }

        private void Add(IEnumerable<CodeModuleMember> codeModuleMembers)
        {
            Add(codeModuleMembers.Select(member => (member is TestCodeModuleMember)
                                                       ? (TestCodeModuleMember) member
                                                       : new TestCodeModuleMember(member)
                    ));
        }

        private void Add(IEnumerable<TestCodeModuleMember> codeModuleMembers)
        {
            _codeModuleMembers.AddRange(codeModuleMembers);
        }

        private static IEnumerable<CodeModuleMember> CreateMembers(IEnumerable<string> methodsUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            //return methodsUnderTest.Select(name => new TestCodeModuleMemberInfo(name, stateUnderTest, expectedBehaviour)).Cast<CodeModuleMemberInfo>().ToList();
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

        public static string GenerateProcedureCode(CodeModuleMember member,
                                                   string stateUnderTest,
                                                   string expectedBehaviour)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, stateUnderTest, expectedBehaviour)))
            {
                return GenerateProcedureCode(new TestCodeModuleMember(member, stateUnderTest, expectedBehaviour));
            }
        }

        internal static string GenerateProcedureCode(TestCodeModuleMember member)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, member.StateUnderTest, member.ExpectedBehaviour)))
            {
                var code = TestMethodTemplate;
                code = code.Replace(MethodUnderTestPlaceholder, member.Name);
                code = code.Replace(StateUnderTestPlaceholder, member.StateUnderTest);
                code = code.Replace(ExpectedBehaviourPlaceholder, member.ExpectedBehaviour);
                code = code.Replace(ParamsPlaceholder, GetProcedureParameterString(member.Name, member.DeclarationString));
                return code;
            }
        }

        private static string GetProcedureParameterString(string procedureName, string procDeclaration)
        {
            if (string.IsNullOrEmpty(procDeclaration))
                return "()"; 

            var parameters = procDeclaration.Substring(procDeclaration.IndexOf(procedureName) + procedureName.Length);
            parameters = ConvertReturnValueToExpected(parameters);
            parameters = ConvertReturnValueToExpectedWithParam(parameters);
            return parameters;
        }

        private static readonly Regex ConvertReturnValueToExpectedRegex = new Regex(@"\(\) As ([^\s]*)", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static string ConvertReturnValueToExpected(string parameters)
        {
            return ConvertReturnValueToExpectedRegex.Replace(parameters,
                                                               m =>
                                                               string.Format("(ByVal Expected As {0})", m.Groups[1].Value));
        }

        private static readonly Regex ConvertReturnValueToExpectedWithParamRegex = new Regex(@"\) As ([^\s]*)", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static string ConvertReturnValueToExpectedWithParam(string parameters)
        {
            return ConvertReturnValueToExpectedWithParamRegex.Replace(parameters,
                                                               m =>
                                                               string.Format(", ByVal Expected As {0})", m.Groups[1].Value));
        }

        internal static string TestClassHeader { get { return TestTemplateSources.TestClassHeader; } }
        private static string TestMethodTemplate { get { return UserSettings.Current.TestMethodTemplate; } }
    }
}