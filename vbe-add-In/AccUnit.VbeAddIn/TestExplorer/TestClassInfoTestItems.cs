namespace AccessCodeLib.AccUnit.VbeAddIn.TestExplorer
{
    public class TestClassInfoTestItems : CheckableItems<TestItem>
    {
        protected override void PerformActionOnAddedItem(TestItem item)
        {
            var testClassInfoTestItem = (TestClassInfoTestItem)item;
            var testClassInfo = testClassInfoTestItem.TestClassInfo;

            if (testClassInfo.Members == null)
                return;

            foreach (var member in testClassInfo.Members)
            {
                var testClassMemberInfoTestItem = new TestClassMemberInfoTestItem(member, true);
                item.Children.Add(testClassMemberInfoTestItem);
            }
        }
    }

    public class TestClassMemberInfoTestItems : CheckableItems<TestItem>
    {
    }

    public class TestRowTestItems : CheckableItems<TestItem>
    {
    }

}
