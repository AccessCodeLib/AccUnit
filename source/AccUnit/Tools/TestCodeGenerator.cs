using AccessCodeLib.AccUnit.Tools.Templates;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        internal static string GenerateTemplateProcedureCode(TestCodeModuleMember member)
        {
            using (new BlockLogger(string.Format(TestMethodNameFormat, member.Name, member.StateUnderTest, member.ExpectedBehaviour)))
            {
                return new TemplateBasedTestMethodBuilder(TemplatesUserSettings.Current.TestMethodTemplate).GenerateProcedureCode(member);
            }
        }

        internal static string TestClassHeader { get { return TestTemplateSources.TestClassHeader; } }
    }
}
