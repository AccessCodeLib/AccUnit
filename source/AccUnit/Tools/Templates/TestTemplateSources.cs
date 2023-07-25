namespace AccessCodeLib.AccUnit.Tools.Templates
{
    static class TestTemplateSources
    {
        private static readonly string TestClassCommonHeader = Properties.Resources.TestClassCommonHeader;
        private static readonly string TestsSectionHeader = Properties.Resources.DefaultTestsSectionHeader;
        internal static readonly string TestClassHeader = TestClassCommonHeader + TestsSectionHeader;
    }
}