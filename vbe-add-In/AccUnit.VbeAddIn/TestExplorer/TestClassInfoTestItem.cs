namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestClassInfoTestItem : TestItem
    {
        public TestClassInfoTestItem(TestClassInfo testClassInfo, bool isChecked = false)
            : base(testClassInfo.Name, testClassInfo.Name, isChecked)
        {
            TestClassInfo = testClassInfo;
        }

        protected override void SetChildren()
        {
            Children = new TestClassMemberInfoTestItems();
        }

        public TestClassInfo TestClassInfo { get; set; }
    }
}
