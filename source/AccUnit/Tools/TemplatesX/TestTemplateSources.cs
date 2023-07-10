using AccessCodeLib.AccUnit.Tools.Resources;

namespace AccessCodeLib.AccUnit.Tools.Templates
{
    static class TestTemplateSources
    {
        private static readonly string TestClassCommonHeader = CodeTemplateParts.TestClassCommonHeader;
        private static readonly string ImplementingInterfaces = CodeTemplateParts.ImplementingInterfaces;
        private static readonly string TestsSectionHeader = CodeTemplateParts.DefaultTestsSectionHeader;
        internal static readonly string TestClassHeaderWithInterfaces = TestClassCommonHeader + ImplementingInterfaces + TestsSectionHeader;
        internal static readonly string TestClassHeader = TestClassCommonHeader + TestsSectionHeader;
    }
}