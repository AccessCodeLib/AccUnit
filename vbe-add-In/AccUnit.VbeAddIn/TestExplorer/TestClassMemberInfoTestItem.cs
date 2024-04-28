namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestClassMemberInfoTestItem : TestItem
    {
        public TestClassMemberInfoTestItem(TestClassMemberInfo testClassMemberInfo, bool isChecked = false)
            : base(testClassMemberInfo.FullName, testClassMemberInfo.Name, isChecked)
        {
            TestClassMemberInfo = testClassMemberInfo;
        }

        protected override void SetChildren()
        {
            Children = new TestRowTestItems();
        }

        public TestClassMemberInfo TestClassMemberInfo { get; set; }
    }
}
