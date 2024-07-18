using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.Tools
{
    public class TestCodeModuleMember : CodeModuleMember
    {
        private const string DefaultStateUnderTestText = @"StateUnderTest";
        private const string DefaultExpectedBehaviourText = @"ExpectedBehaviour";

        public TestCodeModuleMember(CodeModuleMember memberUnderTest,
                                    string stateUnderTest = DefaultStateUnderTestText,
                                    string expectedBehaviour = DefaultExpectedBehaviourText)
            : this(memberUnderTest.Name, memberUnderTest.ProcKind, memberUnderTest.IsPublic, memberUnderTest.DeclarationString, stateUnderTest, expectedBehaviour)
        {
        }

        public TestCodeModuleMember(string methodUnderTest,
                                    string stateUnderTest = DefaultStateUnderTestText,
                                    string expectedBehaviour = DefaultExpectedBehaviourText,
                                    string declarationString = "")
            : this(methodUnderTest, vbext_ProcKind.vbext_pk_Proc, true, declarationString, stateUnderTest, expectedBehaviour)
        {
        }

        private TestCodeModuleMember(string methodUnderTest, vbext_ProcKind procKind, bool isPublic, string declarationString,
                                     string stateUnderTest, string expectedBehaviour)
            : base(methodUnderTest, procKind, isPublic, declarationString)
        {
            using (new BlockLogger())
            {
                StateUnderTest = stateUnderTest;
                ExpectedBehaviour = expectedBehaviour;
            }
        }

        public string StateUnderTest { get; }
        public string ExpectedBehaviour { get; }
    }
}